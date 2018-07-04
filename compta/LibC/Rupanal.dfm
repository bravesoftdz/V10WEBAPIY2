object FRupanal: TFRupanal
  Left = 225
  Top = 158
  HelpContext = 1395200
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Enchainements des ruptures analytiques'
  ClientHeight = 212
  ClientWidth = 337
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PBouton: TPanel
    Left = 0
    Top = 177
    Width = 337
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object BValider: THBitBtn
      Left = 240
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
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
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0127_S16G1'
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 272
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Sortir'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BFermeClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 304
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
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
  end
  object Pappli: TPanel
    Left = 0
    Top = 0
    Width = 337
    Height = 177
    Align = alClient
    TabOrder = 1
    object Ptop: TPanel
      Left = 1
      Top = 1
      Width = 335
      Height = 85
      Align = alTop
      BevelInner = bvRaised
      BevelOuter = bvLowered
      TabOrder = 0
      object TEClibre: THLabel
        Left = 6
        Top = 47
        Width = 70
        Height = 26
        Caption = '&Encha'#238'nement des plans'
        FocusControl = EClibre
        WordWrap = True
      end
      object TCodPlan: THLabel
        Left = 6
        Top = 16
        Width = 25
        Height = 13
        Caption = 'Code'
      end
      object TLibPlan: THLabel
        Left = 137
        Top = 16
        Width = 30
        Height = 13
        Caption = 'Libell'#233
      end
      object EClibre: TEdit
        Left = 81
        Top = 47
        Width = 242
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
      end
      object CodPlan: TEdit
        Left = 81
        Top = 12
        Width = 48
        Height = 21
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
      end
      object LibPlan: TEdit
        Left = 172
        Top = 12
        Width = 151
        Height = 21
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 2
      end
    end
    object Sb1: TScrollBox
      Left = 1
      Top = 86
      Width = 335
      Height = 90
      HorzScrollBar.Visible = False
      Align = alClient
      BorderStyle = bsNone
      Ctl3D = True
      ParentCtl3D = False
      TabOrder = 1
      object TTri1: THLabel
        Left = 4
        Top = 8
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'1'
      end
      object TTri2: THLabel
        Left = 4
        Top = 36
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'2'
      end
      object TTri3: THLabel
        Left = 4
        Top = 64
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'3'
      end
      object TTri4: THLabel
        Left = 4
        Top = 92
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'4'
      end
      object TTri5: THLabel
        Left = 4
        Top = 120
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'5'
      end
      object TTri6: THLabel
        Left = 4
        Top = 147
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'6'
      end
      object TTri7: THLabel
        Left = 4
        Top = 175
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'7'
      end
      object TTri8: THLabel
        Left = 4
        Top = 203
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'8'
      end
      object TTri9: THLabel
        Left = 4
        Top = 231
        Width = 31
        Height = 13
        Caption = 'Tri n'#176'9'
      end
      object TTri10: THLabel
        Left = 4
        Top = 259
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'10'
      end
      object TTri11: THLabel
        Left = 4
        Top = 287
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'11'
      end
      object TTri12: THLabel
        Left = 4
        Top = 315
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'12'
      end
      object TTri13: THLabel
        Left = 4
        Top = 343
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'13'
      end
      object TTri14: THLabel
        Left = 4
        Top = 370
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'14'
      end
      object TTri15: THLabel
        Left = 4
        Top = 398
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'15'
      end
      object TTri16: THLabel
        Left = 4
        Top = 426
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'16'
      end
      object TTri17: THLabel
        Left = 4
        Top = 454
        Width = 37
        Height = 13
        Caption = 'Tri n'#176'17'
      end
      object Tri1: THValComboBox
        Left = 61
        Top = 4
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri2: THValComboBox
        Left = 61
        Top = 32
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri3: THValComboBox
        Left = 61
        Top = 60
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri4: THValComboBox
        Left = 61
        Top = 88
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri5: THValComboBox
        Left = 61
        Top = 116
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri6: THValComboBox
        Left = 61
        Top = 143
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri7: THValComboBox
        Left = 61
        Top = 171
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri8: THValComboBox
        Left = 61
        Top = 199
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri9: THValComboBox
        Left = 61
        Top = 227
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 8
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri10: THValComboBox
        Left = 61
        Top = 255
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 9
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri11: THValComboBox
        Left = 61
        Top = 283
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 10
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri12: THValComboBox
        Left = 61
        Top = 311
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 11
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri13: THValComboBox
        Left = 61
        Top = 339
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 12
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri14: THValComboBox
        Left = 61
        Top = 366
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 13
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri15: THValComboBox
        Left = 61
        Top = 394
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 14
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri16: THValComboBox
        Left = 61
        Top = 422
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 15
        OnClick = Tri1Click
        TagDispatch = 0
      end
      object Tri17: THValComboBox
        Left = 61
        Top = 450
        Width = 250
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 16
        OnClick = Tri1Click
        TagDispatch = 0
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Encha'#238'nement des ruptures analytiques;Un des encha'#238'nements est' +
        ' dupliqu'#233';E;OC;O;O;'
      
        '1;Encha'#238'nement des ruptures analytiques;Un des encha'#238'nements n'#39'e' +
        'st pas un code valide;E;OC;O;O;')
    Left = 80
    Top = 157
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 12
  end
end

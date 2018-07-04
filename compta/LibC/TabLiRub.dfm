object FTabLiRub: TFTabLiRub
  Left = 387
  Top = 88
  HelpContext = 7799100
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Choix sur table libre'
  ClientHeight = 384
  ClientWidth = 424
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
    Top = 349
    Width = 424
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object Panel1: TPanel
      Left = 322
      Top = 2
      Width = 100
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BAide: THBitBtn
        Left = 70
        Top = 2
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
      object BFerme: THBitBtn
        Left = 38
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
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
      object BValider: THBitBtn
        Left = 7
        Top = 2
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
    end
    object CSortia: TEdit
      Left = 135
      Top = 8
      Width = 15
      Height = 21
      CharCase = ecUpperCase
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      Visible = False
    end
    object CEntre: TEdit
      Left = 154
      Top = 8
      Width = 15
      Height = 21
      CharCase = ecUpperCase
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 2
      Visible = False
    end
    object CSorti: TEdit
      Left = 172
      Top = 8
      Width = 15
      Height = 21
      CharCase = ecUpperCase
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 3
      Visible = False
    end
    object Centrea: TEdit
      Left = 191
      Top = 8
      Width = 15
      Height = 21
      CharCase = ecUpperCase
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
      Visible = False
    end
    object CTri: TEdit
      Left = 115
      Top = 7
      Width = 15
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ReadOnly = True
      ShowHint = False
      TabOrder = 5
      Visible = False
    end
    object CbMemo: THValComboBox
      Left = 61
      Top = 5
      Width = 29
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 6
      Visible = False
      TagDispatch = 0
    end
  end
  object PanelChoix: TPanel
    Left = 0
    Top = 0
    Width = 424
    Height = 349
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object LT0: TLabel
      Left = 24
      Top = 14
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 1 de'
    end
    object LT1: TLabel
      Left = 24
      Top = 41
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 2 de'
    end
    object LT2: TLabel
      Left = 24
      Top = 68
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 3 de'
    end
    object LT3: TLabel
      Left = 24
      Top = 96
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 4 de'
    end
    object LT4: TLabel
      Left = 24
      Top = 123
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 5 de'
    end
    object LT5: TLabel
      Left = 24
      Top = 150
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 6 de'
    end
    object LT6: TLabel
      Left = 24
      Top = 177
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 7 de'
    end
    object LT7: TLabel
      Left = 24
      Top = 205
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 8 de'
    end
    object LT8: TLabel
      Left = 24
      Top = 232
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 9 de'
    end
    object LT9: TLabel
      Left = 24
      Top = 259
      Width = 140
      Height = 13
      AutoSize = False
      Caption = 'Table libre 10 de'
    end
    object TComplement: TLabel
      Left = 5
      Top = 291
      Width = 69
      Height = 13
      Caption = '&Et les comptes'
      FocusControl = CompEt
    end
    object Label1: TLabel
      Left = 5
      Top = 319
      Width = 73
      Height = 13
      Caption = '&Ou les comptes'
      FocusControl = CompOu
    end
    object CB0: TCheckBox
      Left = 5
      Top = 12
      Width = 19
      Height = 17
      TabOrder = 0
      OnMouseUp = CB0MouseUp
    end
    object CB1: TCheckBox
      Left = 5
      Top = 39
      Width = 19
      Height = 17
      TabOrder = 1
      OnMouseUp = CB0MouseUp
    end
    object CB2: TCheckBox
      Left = 5
      Top = 66
      Width = 19
      Height = 17
      TabOrder = 2
      OnMouseUp = CB0MouseUp
    end
    object CB3: TCheckBox
      Left = 5
      Top = 94
      Width = 19
      Height = 17
      TabOrder = 3
      OnMouseUp = CB0MouseUp
    end
    object CB4: TCheckBox
      Left = 5
      Top = 121
      Width = 19
      Height = 17
      TabOrder = 4
      OnMouseUp = CB0MouseUp
    end
    object PanelCombos1: TPanel
      Left = 165
      Top = 8
      Width = 257
      Height = 280
      BevelOuter = bvNone
      TabOrder = 5
      object LT00: TLabel
        Left = 125
        Top = 5
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT11: TLabel
        Left = 125
        Top = 32
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT22: TLabel
        Left = 125
        Top = 60
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT33: TLabel
        Left = 125
        Top = 87
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT44: TLabel
        Left = 125
        Top = 114
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT55: TLabel
        Left = 125
        Top = 142
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT66: TLabel
        Left = 125
        Top = 169
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT77: TLabel
        Left = 125
        Top = 196
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT88: TLabel
        Left = 125
        Top = 224
        Width = 6
        Height = 13
        Caption = #224
      end
      object LT99: TLabel
        Left = 125
        Top = 251
        Width = 6
        Height = 13
        Caption = #224
      end
      object T0: THValComboBox
        Left = 1
        Top = 2
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T1: THValComboBox
        Left = 1
        Top = 29
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T2: THValComboBox
        Left = 1
        Top = 57
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T3: THValComboBox
        Left = 1
        Top = 84
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T4: THValComboBox
        Left = 1
        Top = 111
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T00: THValComboBox
        Left = 135
        Top = 2
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T11: THValComboBox
        Left = 135
        Top = 29
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T22: THValComboBox
        Left = 135
        Top = 57
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T33: THValComboBox
        Left = 135
        Top = 84
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 8
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T44: THValComboBox
        Left = 135
        Top = 111
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 9
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T5: THValComboBox
        Left = 1
        Top = 139
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 10
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T6: THValComboBox
        Left = 1
        Top = 166
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 11
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T7: THValComboBox
        Left = 1
        Top = 193
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 12
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T8: THValComboBox
        Left = 1
        Top = 221
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 13
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T9: THValComboBox
        Left = 1
        Top = 248
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 14
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T55: THValComboBox
        Left = 135
        Top = 139
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 15
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T66: THValComboBox
        Left = 135
        Top = 166
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 16
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T77: THValComboBox
        Left = 135
        Top = 193
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 17
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T88: THValComboBox
        Left = 135
        Top = 221
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 18
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T99: THValComboBox
        Left = 135
        Top = 248
        Width = 120
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 19
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
    end
    object CB5: TCheckBox
      Left = 5
      Top = 148
      Width = 18
      Height = 17
      TabOrder = 6
      OnMouseUp = CB0MouseUp
    end
    object CB6: TCheckBox
      Left = 5
      Top = 175
      Width = 18
      Height = 17
      TabOrder = 7
      OnMouseUp = CB0MouseUp
    end
    object CB7: TCheckBox
      Left = 5
      Top = 203
      Width = 18
      Height = 17
      TabOrder = 8
      OnMouseUp = CB0MouseUp
    end
    object CB8: TCheckBox
      Left = 5
      Top = 230
      Width = 18
      Height = 17
      TabOrder = 9
      OnMouseUp = CB0MouseUp
    end
    object CB9: TCheckBox
      Left = 5
      Top = 257
      Width = 18
      Height = 17
      TabOrder = 10
      OnMouseUp = CB0MouseUp
    end
    object CompEt: TEdit
      Left = 87
      Top = 287
      Width = 334
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 11
    end
    object CompOu: TEdit
      Left = 87
      Top = 315
      Width = 334
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 12
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 196
    Top = 72
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Choix sur table libre;La fourchette que vous avez choisie est ' +
        'incoh'#233'rente;W;O;O;O;'
      'Choix sur table libre des comptes g'#233'n'#233'raux'
      'Choix sur table libre des comptes auxiliaires'
      'Choix sur table libre des comptes budg'#233'taires'
      'Choix sur table libre des sections analytiques'
      'Choix sur table libre des sections budg'#233'taires')
    Left = 512
    Top = 56
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Choix sur table libre;Un caract'#232're n'#39'est pas valide dans votre' +
        ' saisie;W;O;O;O;')
    Left = 116
    Top = 76
  end
end

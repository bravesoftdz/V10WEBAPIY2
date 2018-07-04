object FTabLiEdt: TFTabLiEdt
  Left = 180
  Top = 125
  HelpContext = 7490100
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Choix sur table libre'
  ClientHeight = 340
  ClientWidth = 532
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
  object PBouton: TPanel
    Left = 0
    Top = 305
    Width = 532
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object Panel1: TPanel
      Left = 430
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
        GlobalIndexImage = 'Z1770_S16G1'
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
  end
  object PanelTrie: TPanel
    Left = 0
    Top = 61
    Width = 532
    Height = 27
    BevelOuter = bvNone
    TabOrder = 1
    object CTri: TEdit
      Left = 91
      Top = 3
      Width = 437
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ReadOnly = True
      ShowHint = False
      TabOrder = 0
      OnDblClick = CTriDblClick
    end
  end
  object PanelChoix: TPanel
    Left = 0
    Top = 0
    Width = 532
    Height = 305
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object LT0: TLabel
      Left = 34
      Top = 29
      Width = 73
      Height = 13
      Caption = 'Table libre 1 de'
    end
    object LT1: TLabel
      Left = 34
      Top = 56
      Width = 73
      Height = 13
      Caption = 'Table libre 2 de'
    end
    object LT2: TLabel
      Left = 34
      Top = 83
      Width = 73
      Height = 13
      Caption = 'Table libre 3 de'
    end
    object LT3: TLabel
      Left = 34
      Top = 111
      Width = 73
      Height = 13
      Caption = 'Table libre 4 de'
    end
    object LT4: TLabel
      Left = 34
      Top = 138
      Width = 73
      Height = 13
      Caption = 'Table libre 5 de'
    end
    object Label1: TLabel
      Left = 4
      Top = 1
      Width = 117
      Height = 13
      Caption = 'Utiliser les ruptures sur ...'
    end
    object LT5: TLabel
      Left = 34
      Top = 165
      Width = 73
      Height = 13
      Caption = 'Table libre 6 de'
    end
    object LT6: TLabel
      Left = 34
      Top = 192
      Width = 73
      Height = 13
      Caption = 'Table libre 7 de'
    end
    object LT7: TLabel
      Left = 34
      Top = 220
      Width = 73
      Height = 13
      Caption = 'Table libre 8 de'
    end
    object LT8: TLabel
      Left = 34
      Top = 247
      Width = 73
      Height = 13
      Caption = 'Table libre 9 de'
    end
    object LT9: TLabel
      Left = 34
      Top = 274
      Width = 79
      Height = 13
      Caption = 'Table libre 10 de'
    end
    object CB0: TCheckBox
      Left = 5
      Top = 27
      Width = 27
      Height = 17
      TabOrder = 0
      OnMouseUp = CB0MouseUp
    end
    object CB1: TCheckBox
      Left = 5
      Top = 54
      Width = 27
      Height = 17
      TabOrder = 1
      OnMouseUp = CB0MouseUp
    end
    object CB2: TCheckBox
      Left = 5
      Top = 81
      Width = 27
      Height = 17
      TabOrder = 2
      OnMouseUp = CB0MouseUp
    end
    object CB3: TCheckBox
      Left = 5
      Top = 109
      Width = 27
      Height = 17
      TabOrder = 3
      OnMouseUp = CB0MouseUp
    end
    object CB4: TCheckBox
      Left = 5
      Top = 136
      Width = 27
      Height = 17
      TabOrder = 4
      OnMouseUp = CB0MouseUp
    end
    object CbMemo: THValComboBox
      Left = 372
      Top = 8
      Width = 29
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 5
      Visible = False
      TagDispatch = 0
    end
    object PanelCombos1: TPanel
      Left = 212
      Top = 2
      Width = 318
      Height = 294
      BevelOuter = bvNone
      TabOrder = 6
      object LT00: TLabel
        Left = 164
        Top = 27
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T00
      end
      object LT11: TLabel
        Left = 164
        Top = 54
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T11
      end
      object LT22: TLabel
        Left = 164
        Top = 82
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T22
      end
      object LT33: TLabel
        Left = 164
        Top = 109
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T33
      end
      object LT44: TLabel
        Left = 164
        Top = 136
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T44
      end
      object LT55: TLabel
        Left = 164
        Top = 164
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T55
      end
      object LT66: TLabel
        Left = 164
        Top = 191
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T66
      end
      object LT77: TLabel
        Left = 164
        Top = 218
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T77
      end
      object LT88: TLabel
        Left = 164
        Top = 246
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T88
      end
      object LT99: TLabel
        Left = 164
        Top = 273
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = T99
      end
      object LLT0: TLabel
        Left = 2
        Top = 27
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT1: TLabel
        Left = 2
        Top = 54
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT2: TLabel
        Left = 2
        Top = 82
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT3: TLabel
        Left = 2
        Top = 109
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT4: TLabel
        Left = 2
        Top = 136
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT5: TLabel
        Left = 2
        Top = 164
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT6: TLabel
        Left = 2
        Top = 191
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT7: TLabel
        Left = 2
        Top = 218
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT8: TLabel
        Left = 2
        Top = 246
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object LLT9: TLabel
        Left = 2
        Top = 273
        Width = 12
        Height = 13
        Caption = 'de'
        FocusControl = T00
      end
      object T0: THValComboBox
        Left = 17
        Top = 24
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T1: THValComboBox
        Left = 17
        Top = 51
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T2: THValComboBox
        Left = 17
        Top = 79
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T3: THValComboBox
        Left = 17
        Top = 106
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T4: THValComboBox
        Left = 17
        Top = 133
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T00: THValComboBox
        Left = 173
        Top = 24
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T11: THValComboBox
        Left = 173
        Top = 51
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T22: THValComboBox
        Left = 173
        Top = 79
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T33: THValComboBox
        Left = 173
        Top = 106
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 8
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T44: THValComboBox
        Left = 173
        Top = 133
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 9
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T5: THValComboBox
        Left = 17
        Top = 161
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 10
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T6: THValComboBox
        Left = 17
        Top = 188
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 11
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T7: THValComboBox
        Left = 17
        Top = 215
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 12
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T8: THValComboBox
        Left = 17
        Top = 243
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 13
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T9: THValComboBox
        Left = 17
        Top = 270
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 14
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T55: THValComboBox
        Left = 173
        Top = 161
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 15
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T66: THValComboBox
        Left = 173
        Top = 188
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 16
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T77: THValComboBox
        Left = 173
        Top = 215
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 17
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T88: THValComboBox
        Left = 173
        Top = 243
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 18
        OnChange = T0Change
        OnClick = T0Click
        TagDispatch = 0
      end
      object T99: THValComboBox
        Left = 173
        Top = 270
        Width = 145
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
      Top = 163
      Width = 26
      Height = 17
      TabOrder = 7
      OnMouseUp = CB0MouseUp
    end
    object CB6: TCheckBox
      Left = 5
      Top = 190
      Width = 26
      Height = 17
      TabOrder = 8
      OnMouseUp = CB0MouseUp
    end
    object CB7: TCheckBox
      Left = 5
      Top = 218
      Width = 26
      Height = 17
      TabOrder = 9
      OnMouseUp = CB0MouseUp
    end
    object CB8: TCheckBox
      Left = 5
      Top = 245
      Width = 26
      Height = 17
      TabOrder = 10
      OnMouseUp = CB0MouseUp
    end
    object CB9: TCheckBox
      Left = 5
      Top = 272
      Width = 26
      Height = 17
      TabOrder = 11
      OnMouseUp = CB0MouseUp
    end
  end
  object HMTrad: THSystemMenu
    ActiveResize = False
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 144
    Top = 48
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Choix sur tables libres;La fourchette que vous avez choisie es' +
        't incoh'#233'rente;W;O;O;O;'
      'Choix des tables libres des comptes g'#233'n'#233'raux'
      'Choix des tables libres des comptes auxiliaires'
      'Choix des tables libres des comptes budg'#233'taires'
      'Choix des tables libres des sections analytiques'
      'Choix des tables libres des sections budg'#233'taires'
      'Choix des tables libres des immobilisations')
    Left = 300
    Top = 84
  end
end

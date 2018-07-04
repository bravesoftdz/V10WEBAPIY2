object FChoixSSec: TFChoixSSec
  Left = 461
  Top = 221
  ActiveControl = SC1
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Choix d'#39'une section '#224' partir des plans de sous-sections'
  ClientHeight = 133
  ClientWidth = 519
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PBouton: TPanel
    Left = 0
    Top = 98
    Width = 519
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object Panel1: TPanel
      Left = 417
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
        ModalResult = 2
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
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 519
    Height = 98
    Align = alClient
    TabOrder = 1
    object TSC1: TLabel
      Left = 10
      Top = 11
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC1
    end
    object LibSC1: THLabel
      Left = 148
      Top = 11
      Width = 101
      Height = 13
      AutoSize = False
      Caption = 'ttttttttttttttttt'
    end
    object TSC2: TLabel
      Left = 257
      Top = 11
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC2
    end
    object LibSC2: THLabel
      Left = 400
      Top = 11
      Width = 101
      Height = 13
      AutoSize = False
      Caption = 'ttttttttttttttttt'
    end
    object TSC3: TLabel
      Left = 10
      Top = 43
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC3
    end
    object LibSC3: THLabel
      Left = 148
      Top = 43
      Width = 101
      Height = 13
      AutoSize = False
      Caption = 'ttttttttttttttttt'
    end
    object TSC4: TLabel
      Left = 257
      Top = 43
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC4
    end
    object LibSC4: THLabel
      Left = 400
      Top = 43
      Width = 101
      Height = 13
      AutoSize = False
      Caption = 'ttttttttttttttttt'
    end
    object TSC5: TLabel
      Left = 10
      Top = 74
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC5
    end
    object LibSC5: THLabel
      Left = 148
      Top = 74
      Width = 101
      Height = 13
      AutoSize = False
      Caption = 'ttttttttttttttttt'
    end
    object TSC6: TLabel
      Left = 257
      Top = 74
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC6
    end
    object LibSC6: THLabel
      Left = 400
      Top = 74
      Width = 101
      Height = 13
      AutoSize = False
      Caption = 'ttttttttttttttttt'
    end
    object SC1: THCpteEdit
      Left = 65
      Top = 7
      Width = 74
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = LibSC1
      okLocate = True
      SynJoker = False
    end
    object SC2: THCpteEdit
      Left = 316
      Top = 7
      Width = 74
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = LibSC2
      okLocate = True
      SynJoker = False
    end
    object SC3: THCpteEdit
      Left = 65
      Top = 39
      Width = 74
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 2
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = LibSC3
      okLocate = True
      SynJoker = False
    end
    object SC4: THCpteEdit
      Left = 316
      Top = 39
      Width = 74
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 3
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = LibSC4
      okLocate = True
      SynJoker = False
    end
    object SC5: THCpteEdit
      Left = 65
      Top = 70
      Width = 74
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 4
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = LibSC5
      okLocate = True
      SynJoker = False
    end
    object SC6: THCpteEdit
      Left = 316
      Top = 70
      Width = 74
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 5
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = LibSC6
      okLocate = True
      SynJoker = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 9
    Top = 91
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Section analytique;Le code section que vous avez saisi n'#39'exist' +
        'e pas. Voulez-vous le cr'#233'er?;Q;YNC;N;C;')
    Left = 52
    Top = 91
  end
end

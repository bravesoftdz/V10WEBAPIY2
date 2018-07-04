object FZoomSP: TFZoomSP
  Left = 242
  Top = 164
  ActiveControl = SC1
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Choix des sous plans d'#39'une cat'#233'gorie budg'#233'taire'
  ClientHeight = 242
  ClientWidth = 353
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
    Top = 207
    Width = 353
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 2
    object Panel1: TPanel
      Left = 251
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
        Left = 6
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
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 33
    Width = 353
    Height = 174
    Align = alClient
    TabOrder = 1
    object TSC1: TLabel
      Left = 14
      Top = 13
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC1
    end
    object LibSC1: THLabel
      Left = 171
      Top = 13
      Width = 51
      Height = 13
      Caption = 'ttttttttttttttttt'
    end
    object TSC2: TLabel
      Left = 14
      Top = 40
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC2
    end
    object LibSC2: THLabel
      Left = 172
      Top = 40
      Width = 51
      Height = 13
      Caption = 'ttttttttttttttttt'
    end
    object TSC3: TLabel
      Left = 14
      Top = 67
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC3
    end
    object LibSC3: THLabel
      Left = 171
      Top = 67
      Width = 51
      Height = 13
      Caption = 'ttttttttttttttttt'
    end
    object TSC4: TLabel
      Left = 14
      Top = 94
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC4
    end
    object LibSC4: THLabel
      Left = 172
      Top = 94
      Width = 51
      Height = 13
      Caption = 'ttttttttttttttttt'
    end
    object TSC5: TLabel
      Left = 14
      Top = 121
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC5
    end
    object LibSC5: THLabel
      Left = 171
      Top = 121
      Width = 51
      Height = 13
      Caption = 'ttttttttttttttttt'
    end
    object TSC6: TLabel
      Left = 14
      Top = 148
      Width = 27
      Height = 13
      Caption = 'TSC1'
      FocusControl = SC6
    end
    object LibSC6: THLabel
      Left = 172
      Top = 148
      Width = 51
      Height = 13
      Caption = 'ttttttttttttttttt'
    end
    object SC1: THCpteEdit
      Left = 106
      Top = 9
      Width = 53
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
      Left = 106
      Top = 36
      Width = 53
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
      Left = 106
      Top = 63
      Width = 53
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
      Left = 106
      Top = 90
      Width = 53
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
      Left = 106
      Top = 117
      Width = 53
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
      Left = 106
      Top = 144
      Width = 53
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
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 353
    Height = 33
    Align = alTop
    TabOrder = 0
    object FCat: THValComboBox
      Left = 105
      Top = 6
      Width = 189
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FCatChange
      TagDispatch = 0
      DataType = 'TTCATJALBUD'
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 305
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
        'e pas. Voulez-vous le cr'#233'er?;Q;YNC;N;C;'
      'Choix des sous plans d'#39'une cat'#233'gorie budg'#233'taire'
      'Choix des sous plans de la structure analytique')
    Left = 300
    Top = 143
  end
end

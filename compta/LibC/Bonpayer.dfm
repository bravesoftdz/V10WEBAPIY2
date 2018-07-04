object FBonAPayer: TFBonAPayer
  Left = 438
  Top = 240
  HelpContext = 7502000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Emission de bons '#224' payer'
  ClientHeight = 160
  ClientWidth = 271
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
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 0
    Width = 271
    Height = 124
    Align = alClient
  end
  object HLabel1: THLabel
    Left = 14
    Top = 43
    Width = 100
    Height = 13
    AutoSize = False
    Caption = '&Etape du circuit'
    FocusControl = E_SUIVDEC
  end
  object HLabel2: THLabel
    Left = 14
    Top = 70
    Width = 100
    Height = 13
    AutoSize = False
    Caption = '&Nom de lot'
    FocusControl = E_NOMLOT
  end
  object HLabel3: THLabel
    Left = 14
    Top = 97
    Width = 100
    Height = 13
    AutoSize = False
    Caption = '&Mod'#232'le de document'
  end
  object HLabel4: THLabel
    Left = 14
    Top = 16
    Width = 100
    Height = 13
    AutoSize = False
    Caption = 'Compte &s'#233'lection'
    FocusControl = E_GENERAL
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object bListeLots: TToolbarButton97
    Left = 241
    Top = 66
    Width = 16
    Height = 21
    Hint = 'Parcourir'
    Caption = '...'
    ParentShowHint = False
    ShowHint = True
    OnClick = bListeLotsClick
  end
  object Dock: TDock97
    Left = 0
    Top = 124
    Width = 271
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 271
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 271
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 174
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
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
        GlyphMask.Data = {00000000}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAnnuler: TToolbarButton97
        Left = 206
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object Baide: TToolbarButton97
        Left = 238
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BaideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object cModeleBAP: THValComboBox
    Left = 120
    Top = 93
    Width = 138
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 3
    TagDispatch = 0
    DataType = 'TTMODELEBAP'
  end
  object E_SUIVDEC: THValComboBox
    Left = 120
    Top = 39
    Width = 138
    Height = 21
    Style = csDropDownList
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 1
    TagDispatch = 0
    DataType = 'CPCIRCUITUTIL'
  end
  object E_NOMLOT: TEdit
    Left = 120
    Top = 66
    Width = 122
    Height = 21
    CharCase = ecUpperCase
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
  end
  object E_GENERAL: THCpteEdit
    Left = 120
    Top = 12
    Width = 138
    Height = 21
    CharCase = ecUpperCase
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MaxLength = 17
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 0
    ZoomTable = tzGDecais
    Vide = True
    Bourre = False
    okLocate = False
    SynJoker = False
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 174
    Top = 22
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Vous devez renseigner l'#39#233'tape du circuit de validati' +
        'on.;W;O;O;O;'
      '1;?caption?;Vous devez renseigner le nom du lot.;W;O;O;O;'
      
        '2;?caption?;Vous devez choisir un mod'#232'le de bon '#224' payer.;W;O;O;O' +
        ';'
      'Choix d'#39'un nom de lot'
      
        '4;?caption?;Vous devez renseigner un compte de s'#233'lection valide;' +
        'W;O;O;O;')
    Left = 134
    Top = 22
  end
end

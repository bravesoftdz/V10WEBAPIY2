object FParaSpBu: TFParaSpBu
  Left = 211
  Top = 112
  HelpContext = 7577250
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Sous plan par cat'#233'gorie'
  ClientHeight = 381
  ClientWidth = 405
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object TSectBud: TLabel
    Left = 1
    Top = 89
    Width = 200
    Height = 32
    Alignment = taCenter
    AutoSize = False
    Caption = 'Sous plans de l'#39'axe utilis'#233's pour les sections budg'#233'taires'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    WordWrap = True
  end
  object TJalBud: TLabel
    Left = 204
    Top = 89
    Width = 200
    Height = 32
    Alignment = taCenter
    AutoSize = False
    Caption = 'Sous plans de l'#39'axe utilis'#233's pour les journaux budg'#233'taires'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    WordWrap = True
  end
  object HPB: TPanel
    Left = 0
    Top = 346
    Width = 405
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
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
    object Panel1: TPanel
      Left = 302
      Top = 2
      Width = 101
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BAide: THBitBtn
        Left = 68
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
        TabOrder = 0
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: THBitBtn
        Left = 5
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
        TabOrder = 2
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
    Width = 405
    Height = 86
    Align = alTop
    BevelInner = bvLowered
    TabOrder = 1
    object TCC_CODE: TLabel
      Left = 17
      Top = 22
      Width = 25
      Height = 13
      Caption = 'Code'
    end
    object TCC_LIBELLE: TLabel
      Left = 17
      Top = 55
      Width = 30
      Height = 13
      Caption = 'Libell'#233
    end
    object TCbAxe: TLabel
      Left = 185
      Top = 22
      Width = 18
      Height = 13
      Caption = '&Axe'
      FocusControl = CbAxe
    end
    object Cod: TEdit
      Left = 54
      Top = 18
      Width = 46
      Height = 21
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
    end
    object Lib: TEdit
      Left = 54
      Top = 51
      Width = 334
      Height = 21
      ReadOnly = True
      TabOrder = 1
    end
    object CbAxe: THValComboBox
      Left = 220
      Top = 18
      Width = 168
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      OnChange = CbAxeChange
      TagDispatch = 0
      DataType = 'TTAXE'
    end
  end
  object TvSect: TTreeView
    Left = 1
    Top = 151
    Width = 200
    Height = 191
    Images = Image
    Indent = 19
    ReadOnly = True
    TabOrder = 2
    OnChange = TvSectChange
    OnDblClick = TvSectDblClick
  end
  object TvJal: TTreeView
    Left = 204
    Top = 151
    Width = 200
    Height = 191
    Images = Image
    Indent = 19
    ReadOnly = True
    TabOrder = 3
    OnChange = TvSectChange
    OnDblClick = TvSectDblClick
  end
  object SectBud: TEdit
    Left = 1
    Top = 125
    Width = 178
    Height = 21
    ReadOnly = True
    TabOrder = 4
  end
  object JalBud: TEdit
    Left = 204
    Top = 125
    Width = 178
    Height = 21
    ReadOnly = True
    TabOrder = 5
  end
  object BRazSec: THBitBtn
    Left = 180
    Top = 125
    Width = 21
    Height = 21
    Hint = 'Remise '#224' z'#233'ro de la zone'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    OnClick = BRazSecClick
    Margin = 1
    Spacing = -1
    GlobalIndexImage = 'Z0075_S16G1'
    IsControl = True
  end
  object BRazJal: THBitBtn
    Left = 383
    Top = 125
    Width = 21
    Height = 21
    Hint = 'Remise '#224' z'#233'ro de la zone'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    OnClick = BRazJalClick
    Margin = 1
    Spacing = -1
    GlobalIndexImage = 'Z0075_S16G1'
    IsControl = True
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Plan des sous sections'
      'Sous plan par cat'#233'gorie :'
      
        '2;Sous plan par cat'#233'gorie;Voulez-vous enregistrer les modificati' +
        'ons?;Q;YNC;Y;C;'
      
        '3;Sous plan par cat'#233'gorie;Des sous plans sont dupliqu'#233's.Vous dev' +
        'ez reparam'#233'trer la zone.;W;O;O;O;'
      
        '4;Sous plan par cat'#233'gorie;Le nombre de sous plan est sup'#233'rieur a' +
        'u maximum autoris'#233'. Maximun : ;W;O;O;O;')
    Left = 136
    Top = 200
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 72
    Top = 200
  end
  object Image: THImageList
    GlobalIndexImages.Strings = (
      'M0010_S16G1'
      'Z0166_S16G1')
    Left = 262
    Top = 196
  end
end

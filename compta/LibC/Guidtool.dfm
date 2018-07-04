object FGuidTool: TFGuidTool
  Left = 324
  Top = 225
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Choix d'#39'un champ pour les zones du guide'
  ClientHeight = 223
  ClientWidth = 436
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object H_Table: THLabel
    Left = 4
    Top = 12
    Width = 27
    Height = 13
    Caption = '&Table'
    FocusControl = FTable
  end
  object H_NumL: THLabel
    Left = 128
    Top = 61
    Width = 37
    Height = 13
    Caption = '&N'#176' ligne'
    Enabled = False
    FocusControl = FNumL
  end
  object H_Champ: THLabel
    Left = 4
    Top = 38
    Width = 44
    Height = 13
    Caption = '&S'#233'lection'
    FocusControl = FChamp
  end
  object GDate: TGroupBox
    Left = 4
    Top = 111
    Width = 213
    Height = 73
    Caption = ' Dates '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    object H_Format: THLabel
      Left = 8
      Top = 20
      Width = 32
      Height = 13
      Caption = 'Format'
      FocusControl = FFormat
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object H_ExempleD: THLabel
      Left = 8
      Top = 48
      Width = 40
      Height = 13
      Caption = 'Exemple'
      FocusControl = FExempleD
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object FFormat: TComboBox
      Left = 52
      Top = 16
      Width = 153
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
      OnChange = FFormatChange
      Items.Strings = (
        'dd/mm/yy'
        'dd/mm/yyyy'
        'dd-mmm-yy'
        'dd-mmm-yyyy'
        'dd mmm yyyy')
    end
    object FExempleD: TEdit
      Left = 52
      Top = 44
      Width = 153
      Height = 21
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
  end
  object GString: TGroupBox
    Left = 4
    Top = 113
    Width = 213
    Height = 73
    Caption = ' Cha'#238'nes '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    object H_Debut: THLabel
      Left = 8
      Top = 20
      Width = 29
      Height = 13
      Caption = '&D'#233'but'
      FocusControl = FDebut
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object H_Longueur: THLabel
      Left = 116
      Top = 20
      Width = 45
      Height = 13
      Caption = '&Longueur'
      FocusControl = FLongueur
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object H_exempleS: THLabel
      Left = 8
      Top = 48
      Width = 40
      Height = 13
      Caption = '&Exemple'
      FocusControl = FExempleS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object FDebut: TSpinEdit
      Left = 52
      Top = 15
      Width = 37
      Height = 22
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxValue = 200
      MinValue = 1
      ParentFont = False
      TabOrder = 0
      Value = 1
      OnChange = FDebutChange
    end
    object FLongueur: TSpinEdit
      Left = 168
      Top = 15
      Width = 37
      Height = 22
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxValue = 200
      MinValue = 1
      ParentFont = False
      TabOrder = 1
      Value = 1
      OnChange = FDebutChange
    end
    object FExempleS: TEdit
      Left = 52
      Top = 44
      Width = 153
      Height = 21
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 2
    end
  end
  object FTable: THValComboBox
    Left = 53
    Top = 8
    Width = 160
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 0
    OnChange = FTableChange
    TagDispatch = 0
  end
  object FListe: TListBox
    Left = 219
    Top = 4
    Width = 214
    Height = 182
    ItemHeight = 13
    TabOrder = 1
    OnClick = FListeClick
    OnDblClick = FListeDblClick
  end
  object FChamp: TEdit
    Left = 53
    Top = 35
    Width = 160
    Height = 21
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
  end
  object FNumL: TSpinEdit
    Left = 173
    Top = 58
    Width = 37
    Height = 22
    Enabled = False
    MaxValue = 200
    MinValue = 1
    TabOrder = 3
    Value = 1
  end
  object PBouton: TPanel
    Left = 0
    Top = 188
    Width = 436
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 7
    object BValider: THBitBtn
      Left = 335
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider le choix'
      Default = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValiderClick
      NumGlyphs = 2
      GlobalIndexImage = 'Z0003_S16G2'
    end
    object BFermer: THBitBtn
      Left = 367
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BFermerClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 399
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
      Layout = blGlyphRight
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
  end
  object FCache: THValComboBox
    Left = 32
    Top = 176
    Width = 49
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 6
    Visible = False
    OnChange = FTableChange
    TagDispatch = 0
  end
  object CLigCur: TRadioButton
    Left = 4
    Top = 77
    Width = 113
    Height = 17
    Alignment = taLeftJustify
    Caption = 'Ligne &courante'
    Checked = True
    TabOrder = 8
    TabStop = True
    OnClick = CLigFixeClick
  end
  object CLigPrec: TRadioButton
    Left = 4
    Top = 94
    Width = 113
    Height = 17
    Alignment = taLeftJustify
    Caption = 'Ligne &pr'#233'c'#233'dente'
    TabOrder = 9
    OnClick = CLigFixeClick
  end
  object CLigFixe: TRadioButton
    Left = 4
    Top = 61
    Width = 113
    Height = 17
    Alignment = taLeftJustify
    Caption = 'Ligne &fixe'
    TabOrder = 10
    OnClick = CLigFixeClick
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 116
    Top = 16
  end
end

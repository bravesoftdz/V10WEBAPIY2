object FZoom: TFZoom
  Left = 313
  Top = 202
  HelpContext = 1300
  HorzScrollBar.Visible = False
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Recherche d'#39'un'
  ClientHeight = 295
  ClientWidth = 352
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel1: TPanel
    Left = 0
    Top = 0
    Width = 352
    Height = 41
    Align = alTop
    Caption = ' '
    TabOrder = 0
    object TCompte: TLabel
      Left = 9
      Top = 12
      Width = 36
      Height = 13
      Caption = 'Compte'
      FocusControl = SelectCompte
      OnClick = SelectCompteChange
      OnMouseDown = TCompteMouseDown
    end
    object TLibelle: TLabel
      Left = 157
      Top = 12
      Width = 30
      Height = 13
      Caption = 'Libell'#233
      FocusControl = Selectlib
      OnClick = SelectCompteChange
      OnMouseDown = TCompteMouseDown
    end
    object SelectCompte: TEdit
      Left = 56
      Top = 8
      Width = 85
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 17
      TabOrder = 0
      Text = 'NE PAS VIRER'
      OnChange = SelectCompteChange
    end
    object Selectlib: TEdit
      Tag = 1
      Left = 196
      Top = 8
      Width = 129
      Height = 21
      MaxLength = 35
      TabOrder = 1
      OnChange = SelectCompteChange
    end
  end
  object HPanel2: TPanel
    Left = 0
    Top = 259
    Width = 352
    Height = 36
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    TabOrder = 2
    object OKBtn: THBitBtn
      Left = 256
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Ouvrir'
      Default = True
      ModalResult = 1
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object CancelBtn: THBitBtn
      Left = 288
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object HelpBtn: THBitBtn
      Left = 320
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = HelpBtnClick
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
        A600000000000000000000000000000000000000000000000000000000000000
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
        03030606030303030303030303030303030303FFFF0303030303030303030303
        0303030303060404060303030303030303030303030303F8F8FF030303030303
        030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
        0303030303030303030303030303FE060603030303030303030303030303F8FF
        FFF8FF0303030303030303030303030303030303030303030303030303030303
        030303F8F8030303030303030303030303030303030304040603030303030303
        0303030303030303FFFF03030303030303030303030303030306060604030303
        0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
        0403030303030303030303030303F8FF03F8FF03030303030303030303030303
        03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
        030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
        0303030303030303030303FE060606040303030303030303030303F8FF0303F8
        FF030303030303030303030404030303FE060606040303030303030303FF0303
        F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
        F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
        030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
        06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
        060606060303030303030303F8FFFF030303030303F803030303030303030303
        0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
        0303030303030303030303030303030303030303030303F8F8F8F8F803030303
        0303}
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object Appelf: THBitBtn
      Left = 4
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Zoom'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = APPELFICHE
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
    object CreatBn: THBitBtn
      Left = 36
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Nouveau'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = CREATEFICHE
      GlobalIndexImage = 'Z0053_S16G1'
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 41
    Width = 352
    Height = 218
    Align = alClient
    ColCount = 3
    DefaultColWidth = 128
    DefaultRowHeight = 18
    TabOrder = 1
    OnDblClick = FListeDblClick
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    DBIndicator = True
    ColWidths = (
      10
      128
      189)
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Recherche d'#39'un compte g'#233'n'#233'ral'
      'Recherche d'#39'un compte auxiliaire'
      'Recherche d'#39'une section analytique'
      'Recherche d'#39'un journal'
      'Recherche d'#39'un article'
      'Recherche d'#39'un compte budg'#233'taire'
      'Recherche d'#39'une immobilisation'
      'Recherche d'#39'une valeur'
      'Compte'
      'Auxiliaire'
      'Section'
      'Journal'
      'Article'
      'Compte'
      'Compte'
      'Libell'#233
      'Recherche d'#39'un compte de correspondance'
      'Recherche d'#39'une nomenclature'
      'Nomenclature'
      'Code'
      'Recherche d'#39'un section budg'#233'taire'
      'Recherche d'#39'un journal budg'#233'taire'
      'Section'
      'Journal'
      'Sous sect.'
      'Recherche d'#39'une sous section'
      'Code'
      'Rubrique'
      'Recherche d'#39'une rubrique'
      'Recherche de'
      '30;')
    Left = 136
    Top = 164
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 132
    Top = 92
  end
end

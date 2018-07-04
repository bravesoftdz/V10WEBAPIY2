object FSaisEcar: TFSaisEcar
  Left = 256
  Top = 163
  HelpContext = 7244600
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSingle
  Caption = 'R'#233'gularisation d'#39#233'carts de conversion en saisie'
  ClientHeight = 348
  ClientWidth = 612
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object POutils: TPanel
    Left = 0
    Top = 315
    Width = 612
    Height = 33
    Align = alBottom
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 0
    object BValide: THBitBtn
      Left = 512
      Top = 3
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
      OnClick = BValideClick
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
    object BAbandon: THBitBtn
      Left = 544
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BAbandonClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BAide: THBitBtn
      Left = 576
      Top = 3
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
    object BSolde: THBitBtn
      Left = 4
      Top = 3
      Width = 28
      Height = 27
      Hint = 'R'#233'percuter l'#39#233'cart sur la ligne courante'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BSoldeClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0051_S16G2'
      IsControl = True
    end
    object BModif: THBitBtn
      Left = 36
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Modifier manuellement les '#233'carts'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = BModifClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1334_S16G1'
      IsControl = True
    end
  end
  object PEntete: TPanel
    Left = 0
    Top = 288
    Width = 612
    Height = 27
    Align = alBottom
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 1
    object HOppose: THLabel
      Left = 421
      Top = 7
      Width = 33
      Height = 13
      Alignment = taRightJustify
      Caption = 'Solde'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Bevel4: TBevel
      Left = 458
      Top = 4
      Width = 116
      Height = 18
    end
    object HPivot: THLabel
      Left = 215
      Top = 7
      Width = 33
      Height = 13
      Alignment = taRightJustify
      Caption = 'Solde'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Bevel1: TBevel
      Left = 252
      Top = 4
      Width = 116
      Height = 18
    end
    object SA_SOLDEOppose: THNumEdit
      Tag = 1
      Left = 459
      Top = 6
      Width = 114
      Height = 15
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object SA_SoldePivot: THNumEdit
      Tag = 1
      Left = 253
      Top = 6
      Width = 114
      Height = 15
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
  end
  object PFen: TPanel
    Left = 0
    Top = 0
    Width = 612
    Height = 288
    Align = alClient
    TabOrder = 2
    object LPivot: TLabel
      Left = 211
      Top = 8
      Width = 175
      Height = 21
      Alignment = taCenter
      AutoSize = False
      Caption = 'LPivot'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object LOppose: TLabel
      Left = 387
      Top = 8
      Width = 174
      Height = 21
      Alignment = taCenter
      AutoSize = False
      Caption = 'LOppose'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label1: TLabel
      Left = 48
      Top = 8
      Width = 116
      Height = 16
      Caption = 'Lignes d'#39#233'criture'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object ISigneEuro: TImage
      Left = 16
      Top = 7
      Width = 16
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        0000100000000100040000000000800000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
        FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
        4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
        FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
        4FF4}
      Stretch = True
      Transparent = True
    end
    object G: THGrid
      Left = 3
      Top = 32
      Width = 609
      Height = 249
      ColCount = 9
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goTabs, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Lig'
        'G'#233'n'#233'ral'
        'Auxiliaire'
        'D'#233'bit'
        'Cr'#233'dit'
        'Solde'
        'D'#233'bit'
        'Cr'#233'dit'
        'Solde')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnCellEnter = GCellEnter
      OnCellExit = GCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        18
        54
        60
        74
        74
        79
        76
        73
        74)
      RowHeights = (
        18
        18
        18
        18
        18)
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 72
    Top = 212
  end
  object HDiv: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'EURO'
      
        '1;Ecarts de conversion;Validation impossible, vos soldes ne sont' +
        ' pas nuls ;W;O;O;O;'
      
        '2;Ecarts de conversion;Attention ! L'#39#233'criture n'#39'est pas '#233'quilibr' +
        #233'e, elle ne pourra pas '#234'tre valid'#233'e. Confirmez-vous la sortie ?;' +
        'W;YN;N;N;'
      
        '3;Ecarts de conversion;Les montants th'#233'oriques diff'#232'rent sensibl' +
        'ement de ceux saisis. Confirmez-vous l'#39'op'#233'ration ?;Q;YN;Y;Y;'
      'Visualisation des montants en'
      
        '5;Ecarts de conversion;Les montants saisis et calcul'#233's ne sont p' +
        'as coh'#233'rents. Validation impossible;W;O;O;O;'
      'recalcul'#233's')
    Left = 114
    Top = 213
  end
end

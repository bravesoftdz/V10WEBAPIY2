object FJournal: TFJournal
  Left = 156
  Top = 227
  HelpContext = 7208000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Journal comptable :'
  ClientHeight = 301
  ClientWidth = 468
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPB: TPanel
    Left = 0
    Top = 266
    Width = 468
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    TabOrder = 0
    object FAutoSave: TCheckBox
      Left = 246
      Top = 12
      Width = 27
      Height = 13
      Caption = 'Auto.'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 10
      Visible = False
    end
    object BImprimer: THBitBtn
      Left = 339
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Caption = ' '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BAnnuler: THBitBtn
      Left = 120
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Annuler la derni'#232're action'
      Cancel = True
      Caption = ' '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = BAnnulerClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0075_S16G1'
      IsControl = True
    end
    object BFirst: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Premier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrior: THBitBtn
      Left = 32
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Pr'#233'c'#233'dent'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BPriorClick
      GlobalIndexImage = 'Z0017_S16G1'
    end
    object BNext: THBitBtn
      Left = 60
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Suivant'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BNextClick
      GlobalIndexImage = 'Z0031_S16G1'
    end
    object BLast: THBitBtn
      Left = 88
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Dernier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BLastClick
      GlobalIndexImage = 'Z0264_S16G1'
    end
    object Binsert: THBitBtn
      Left = 152
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Nouveau'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnClick = BinsertClick
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object BFerme: THBitBtn
      Left = 404
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = BFermeClick
      GlobalIndexImage = 'Z1770_S16G1'
    end
    object BValider: THBitBtn
      Left = 372
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
      TabOrder = 7
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
    object BAide: THBitBtn
      Left = 436
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object BMenuZoom: THBitBtn
      Tag = -100
      Left = 185
      Top = 4
      Width = 37
      Height = 27
      Hint = 'Menu zoom'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 11
      OnClick = BMenuZoomClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0016_S16G1'
      IsControl = True
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 468
    Height = 266
    ActivePage = PCar
    Align = alClient
    TabOrder = 1
    OnChanging = PagesChanging
    object PCar: TTabSheet
      Caption = 'Caract'#233'ristiques'
      object TJ_JOURNAL: THLabel
        Left = 3
        Top = 18
        Width = 25
        Height = 13
        Caption = '&Code'
        FocusControl = J_JOURNAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_NATUREJAL: THLabel
        Left = 101
        Top = 18
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = J_NATUREJAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_LIBELLE: THLabel
        Left = 4
        Top = 97
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = J_LIBELLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_ABREGE: THLabel
        Left = 4
        Top = 121
        Width = 34
        Height = 13
        Caption = '&Abr'#233'g'#233
        FocusControl = J_ABREGE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_AXE: THLabel
        Left = 294
        Top = 18
        Width = 18
        Height = 13
        Caption = 'A&xe'
        FocusControl = J_AXE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_MODESAISIE: THLabel
        Left = 4
        Top = 145
        Width = 56
        Height = 13
        Caption = '&Mode saisie'
        FocusControl = J_MODESAISIE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HGBDERNMOUV: TGroupBox
        Left = 2
        Top = 44
        Width = 454
        Height = 41
        Caption = ' Dernier Mouvement '
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        object TJ_DEBITDERNMVT: THLabel
          Left = 288
          Top = -1
          Width = 31
          Height = 13
          Caption = ' D'#233'bit '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_CREDITDERNMVT: THLabel
          Left = 416
          Top = -1
          Width = 33
          Height = 13
          Caption = ' Cr'#233'dit '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_NUMDERNMVT: THLabel
          Left = 137
          Top = -1
          Width = 33
          Height = 13
          Caption = ' Pi'#232'ce '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object J_DEBITDERNMVT: TDBEdit
          Left = 190
          Top = 16
          Width = 126
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_DEBITDERNMVT'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
        end
        object J_CREDITDERNMVT: TDBEdit
          Left = 320
          Top = 16
          Width = 126
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_CREDITDERNMVT'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
        end
        object J_NUMDERNMVT: TDBEdit
          Left = 114
          Top = 16
          Width = 53
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_NUMDERNMVT'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
        object J_DATEDERNMVT: TDBEdit
          Left = 12
          Top = 16
          Width = 75
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_DATEDERNMVT'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 3
        end
      end
      object J_JOURNAL: TDBEdit
        Left = 42
        Top = 14
        Width = 49
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'J_JOURNAL'
        DataSource = SJal
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        OnChange = J_JOURNALChange
        OnEnter = J_JOURNALChange
        OnExit = J_JOURNALExit
      end
      object J_NATUREJAL: THDBValComboBox
        Left = 138
        Top = 14
        Width = 147
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        OnChange = J_NATUREJALChange
        TagDispatch = 0
        DataType = 'TTNATJAL'
        DataField = 'J_NATUREJAL'
        DataSource = SJal
      end
      object J_LIBELLE: TDBEdit
        Left = 65
        Top = 94
        Width = 247
        Height = 21
        Ctl3D = True
        DataField = 'J_LIBELLE'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
      end
      object J_ABREGE: TDBEdit
        Left = 65
        Top = 117
        Width = 119
        Height = 21
        Ctl3D = True
        DataField = 'J_ABREGE'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
      end
      object J_MULTIDEVISE: TDBCheckBox
        Left = 328
        Top = 98
        Width = 116
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Multi-&devises'
        DataField = 'J_MULTIDEVISE'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object J_CENTRALISABLE: TDBCheckBox
        Left = 328
        Top = 121
        Width = 116
        Height = 13
        Alignment = taLeftJustify
        Caption = 'C&entralisable'
        DataField = 'J_CENTRALISABLE'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object J_AXE: THDBValComboBox
        Left = 320
        Top = 14
        Width = 135
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        OnChange = J_NATUREJALChange
        TagDispatch = 0
        DataType = 'TTAXE'
        DataField = 'J_AXE'
        DataSource = SJal
      end
      object GBCarac: TGroupBox
        Left = 1
        Top = 164
        Width = 455
        Height = 70
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        object HLabel1: THLabel
          Left = 162
          Top = 0
          Width = 31
          Height = 13
          Caption = ' D'#233'bit '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel3: THLabel
          Left = 288
          Top = 0
          Width = 33
          Height = 13
          Caption = ' Cr'#233'dit '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel2: THLabel
          Left = 414
          Top = 0
          Width = 33
          Height = 13
          Caption = ' Solde '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel4: THLabel
          Left = 9
          Top = 0
          Width = 52
          Height = 13
          Caption = ' Exercices '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel5: THLabel
          Left = 11
          Top = 16
          Width = 49
          Height = 13
          Caption = 'Pr'#233'c'#233'dent'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clMaroon
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel6: THLabel
          Left = 11
          Top = 32
          Width = 42
          Height = 13
          Caption = 'En cours'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel7: THLabel
          Left = 11
          Top = 48
          Width = 36
          Height = 13
          Caption = 'Suivant'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGreen
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object J_TOTDEBP: TDBEdit
          Left = 64
          Top = 16
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_TOTDEBP'
          DataSource = SJal
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clMaroon
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
        object J_TOTCREP: TDBEdit
          Left = 192
          Top = 16
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_TOTCREP'
          DataSource = SJal
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clMaroon
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
        end
        object JSOLCREP: THNumEdit
          Left = 320
          Top = 16
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clMaroon
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          NumericType = ntDC
          UseRounding = True
          Validate = False
        end
        object J_TOTDEBE: TDBEdit
          Left = 64
          Top = 32
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_TOTDEBE'
          DataSource = SJal
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
        end
        object J_TOTCREE: TDBEdit
          Left = 192
          Top = 32
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_TOTCREE'
          DataSource = SJal
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 4
        end
        object JSOLCREE: THNumEdit
          Left = 320
          Top = 32
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 5
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          NumericType = ntDC
          UseRounding = True
          Validate = False
        end
        object J_TOTDEBS: TDBEdit
          Left = 64
          Top = 48
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_TOTDEBS'
          DataSource = SJal
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGreen
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 6
        end
        object J_TOTCRES: TDBEdit
          Left = 192
          Top = 48
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_TOTCRES'
          DataSource = SJal
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGreen
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 7
        end
        object JSOLCRES: THNumEdit
          Left = 320
          Top = 48
          Width = 126
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGreen
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 8
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          NumericType = ntDC
          UseRounding = True
          Validate = False
        end
      end
      object J_EFFET: TDBCheckBox
        Left = 328
        Top = 147
        Width = 116
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Suivi d'#39'e&ffets'
        Color = clBtnFace
        DataField = 'J_EFFET'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 8
        ValueChecked = 'X'
        ValueUnchecked = '-'
        OnClick = J_EFFETClick
      end
      object J_MODESAISIE: THDBValComboBox
        Left = 65
        Top = 141
        Width = 119
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 5
        OnChange = J_MODESAISIEChange
        TagDispatch = 0
        DataType = 'CPMODESAISIEJAL'
        DataField = 'J_MODESAISIE'
        DataSource = SJal
      end
    end
    object PComplement: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel2: TBevel
        Left = 5
        Top = 10
        Width = 448
        Height = 199
      end
      object TJ_COMPTEINTERDIT: THLabel
        Left = 8
        Top = 24
        Width = 80
        Height = 13
        Caption = 'Comptes &interdits'
        FocusControl = J_COMPTEINTERDIT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_COMPTEAUTOMAT: THLabel
        Left = 8
        Top = 57
        Width = 107
        Height = 13
        Caption = 'Comptes &automatiques'
        FocusControl = J_COMPTEAUTOMAT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_COMPTEURNORMAL: THLabel
        Left = 8
        Top = 89
        Width = 79
        Height = 13
        Caption = '&Compteur normal'
        FocusControl = J_COMPTEURNORMAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_COMPTEURSIMUL: THLabel
        Left = 8
        Top = 121
        Width = 109
        Height = 13
        Caption = 'Compteur de &simulation'
        FocusControl = J_COMPTEURSIMUL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_CONTREPARTIE: THLabel
        Left = 8
        Top = 153
        Width = 110
        Height = 13
        Caption = 'Compte de c&ontrepartie'
        FocusControl = J_CONTREPARTIE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TJ_TYPECONTREPARTIE: THLabel
        Left = 8
        Top = 183
        Width = 98
        Height = 13
        Caption = '&Type de contrepartie'
        FocusControl = J_TYPECONTREPARTIE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object J_COMPTEINTERDIT: TDBEdit
        Left = 124
        Top = 20
        Width = 321
        Height = 21
        Ctl3D = True
        DataField = 'J_COMPTEINTERDIT'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object J_COMPTEAUTOMAT: TDBEdit
        Left = 124
        Top = 52
        Width = 321
        Height = 21
        Ctl3D = True
        DataField = 'J_COMPTEAUTOMAT'
        DataSource = SJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object J_COMPTEURNORMAL: THDBValComboBox
        Left = 124
        Top = 84
        Width = 157
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TTSOUCHECOMPTA'
        DataField = 'J_COMPTEURNORMAL'
        DataSource = SJal
      end
      object J_COMPTEURSIMUL: THDBValComboBox
        Left = 124
        Top = 117
        Width = 157
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTSOUCHECOMPTASIMUL'
        DataField = 'J_COMPTEURSIMUL'
        DataSource = SJal
      end
      object J_CONTREPARTIE: THDBCpteEdit
        Left = 124
        Top = 149
        Width = 157
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        OnExit = J_CPTEREGULDEBIT1Exit
        ZoomTable = tzGbanque
        Vide = False
        Bourre = True
        okLocate = False
        SynJoker = False
        DataField = 'J_CONTREPARTIE'
        DataSource = SJal
      end
      object J_TYPECONTREPARTIE: THDBValComboBox
        Left = 124
        Top = 180
        Width = 157
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTTYPECONTREPARTIE'
        DataField = 'J_TYPECONTREPARTIE'
        DataSource = SJal
      end
    end
    object PInf: TTabSheet
      Caption = 'Informations'
      object TJ_BLOCNOTE: TGroupBox
        Left = 7
        Top = 56
        Width = 446
        Height = 152
        Caption = ' Bloc Notes '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object J_BLOCNOTE: THDBRichEditOLE
          Left = 8
          Top = 16
          Width = 430
          Height = 126
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          MaxLength = 262140
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 0
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1036{\fonttbl{\f' +
              '0\fswiss Arial;}}'
            '{\colortbl ;\red0\green0\blue0;}'
            '{\*\generator Riched20 10.0.14393}\viewkind4\uc1 '
            '\pard\cf1\f0\fs16 J_BLOCNOTE'
            '\par '
            '\par }')
          DataField = 'J_BLOCNOTE'
          DataSource = SJal
        end
      end
      object HGBDates: TGroupBox
        Left = 7
        Top = 6
        Width = 446
        Height = 41
        Caption = ' Dates '
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TJ_DATEOUVERTURE: THLabel
          Left = 163
          Top = 0
          Width = 49
          Height = 13
          Caption = ' Ouvert le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_DATEFERMETURE: THLabel
          Left = 379
          Top = 0
          Width = 46
          Height = 13
          Caption = ' Ferm'#233' le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_DATEMODIFICATION: THLabel
          Left = 273
          Top = 0
          Width = 51
          Height = 13
          Caption = ' Modifi'#233' le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_DATECREATION: THLabel
          Left = 60
          Top = 0
          Width = 39
          Height = 13
          Caption = ' Cr'#233'e le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object J_FERME: TDBCheckBox
          Left = 116
          Top = 32
          Width = 29
          Height = 13
          Caption = 'Ferm'#233
          Color = clYellow
          DataField = 'J_FERME'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
          ValueChecked = 'X'
          ValueUnchecked = '-'
          Visible = False
        end
        object J_DATECREATION: TDBEdit
          Left = 38
          Top = 15
          Width = 75
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_DATECREATION'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 1
        end
        object J_DATEOUVERTURE: TDBEdit
          Left = 151
          Top = 15
          Width = 75
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_DATEOUVERTURE'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 2
        end
        object J_DATEMODIF: TDBEdit
          Left = 264
          Top = 15
          Width = 75
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_DATEMODIF'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 3
        end
        object J_DATEFERMETURE: TDBEdit
          Left = 365
          Top = 15
          Width = 75
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'J_DATEFERMETURE'
          DataSource = SJal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 4
        end
      end
    end
    object PReg: TTabSheet
      Caption = 'R'#233'gularisations'
      object GbReg: TGroupBox
        Left = 5
        Top = 15
        Width = 449
        Height = 173
        Caption = ' Comptes de r'#233'gularisations '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TJ_CPTEREGULDEBIT1: TLabel
          Left = 11
          Top = 47
          Width = 53
          Height = 13
          Caption = '1er compte'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_CPTEREGULDEBIT2: TLabel
          Left = 11
          Top = 86
          Width = 64
          Height = 13
          Caption = '2'#232'me compte'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TJ_CPTEREGULDEBIT3: TLabel
          Left = 11
          Top = 126
          Width = 64
          Height = 13
          Caption = '3'#232'me compte'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label1: TLabel
          Left = 164
          Top = 26
          Width = 25
          Height = 13
          Caption = 'D'#233'bit'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label2: TLabel
          Left = 345
          Top = 26
          Width = 27
          Height = 13
          Caption = 'Cr'#233'dit'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object J_CPTEREGULDEBIT1: THDBCpteEdit
          Left = 101
          Top = 43
          Width = 151
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnExit = J_CPTEREGULDEBIT1Exit
          ZoomTable = tzGNLettColl
          Vide = False
          Bourre = True
          okLocate = False
          SynJoker = False
          DataField = 'J_CPTEREGULDEBIT1'
          DataSource = SJal
        end
        object J_CPTEREGULCREDIT1: THDBCpteEdit
          Left = 283
          Top = 43
          Width = 151
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnExit = J_CPTEREGULDEBIT1Exit
          ZoomTable = tzGNLettColl
          Vide = False
          Bourre = True
          okLocate = False
          SynJoker = False
          DataField = 'J_CPTEREGULCREDIT1'
          DataSource = SJal
        end
        object J_CPTEREGULDEBIT2: THDBCpteEdit
          Left = 101
          Top = 82
          Width = 151
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnExit = J_CPTEREGULDEBIT1Exit
          ZoomTable = tzGNLettColl
          Vide = False
          Bourre = True
          okLocate = False
          SynJoker = False
          DataField = 'J_CPTEREGULDEBIT2'
          DataSource = SJal
        end
        object J_CPTEREGULCREDIT2: THDBCpteEdit
          Left = 283
          Top = 82
          Width = 151
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnExit = J_CPTEREGULDEBIT1Exit
          ZoomTable = tzGNLettColl
          Vide = False
          Bourre = True
          okLocate = False
          SynJoker = False
          DataField = 'J_CPTEREGULCREDIT2'
          DataSource = SJal
        end
        object J_CPTEREGULDEBIT3: THDBCpteEdit
          Left = 101
          Top = 122
          Width = 151
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnExit = J_CPTEREGULDEBIT1Exit
          ZoomTable = tzGNLettColl
          Vide = False
          Bourre = True
          okLocate = False
          SynJoker = False
          DataField = 'J_CPTEREGULDEBIT3'
          DataSource = SJal
        end
        object J_CPTEREGULCREDIT3: THDBCpteEdit
          Left = 283
          Top = 122
          Width = 151
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          OnExit = J_CPTEREGULDEBIT1Exit
          ZoomTable = tzGNLettColl
          Vide = False
          Bourre = True
          okLocate = False
          SynJoker = False
          DataField = 'J_CPTEREGULCREDIT3'
          DataSource = SJal
        end
      end
    end
    object PTreso: TTabSheet
      Caption = 'Saisie de tr'#233'sorerie'
      ImageIndex = 4
      object HLabel9: THLabel
        Left = 8
        Top = 16
        Width = 75
        Height = 13
        Caption = 'Date comptable'
      end
      object J_TRESODATE: THDBValComboBox
        Left = 120
        Top = 12
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        TagDispatch = 0
        DataType = 'CPTRESODATE'
        DataField = 'J_CHOIXDATE'
        DataSource = SJal
      end
    end
    object PSAISIE: TTabSheet
      Caption = 'Saisie '
      ImageIndex = 5
      object HLabel8: THLabel
        Left = 35
        Top = 101
        Width = 83
        Height = 13
        Caption = 'Nature par d'#233'faut'
      end
      object J_INCREF: TDBCheckBox
        Left = 35
        Top = 15
        Width = 172
        Height = 21
        Alignment = taLeftJustify
        Caption = 'Incr'#233'menter la r'#233'f'#233'rence'
        DataField = 'J_INCREF'
        DataSource = SJal
        TabOrder = 0
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object J_NATCOMPL: TDBCheckBox
        Left = 35
        Top = 130
        Width = 172
        Height = 16
        Alignment = taLeftJustify
        Caption = 'Afficher le code de la nature'
        DataField = 'J_NATCOMPL'
        DataSource = SJal
        TabOrder = 1
        ValueChecked = 'X'
        ValueUnchecked = '-'
        Visible = False
      end
      object J_EQAUTO: TDBCheckBox
        Left = 35
        Top = 44
        Width = 172
        Height = 21
        Alignment = taLeftJustify
        Caption = 'Solde automatique'
        DataField = 'J_EQAUTO'
        DataSource = SJal
        TabOrder = 2
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object J_INCNUM: TDBCheckBox
        Left = 35
        Top = 73
        Width = 172
        Height = 21
        Alignment = taLeftJustify
        Caption = 'Incr'#233'menter le bordereau'
        DataField = 'J_INCNUM'
        DataSource = SJal
        TabOrder = 3
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object J_NATDEFAUT: THDBValComboBox
        Left = 195
        Top = 99
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        VideString = 'Aucun'
        DataType = 'TTNATUREPIECE'
        DataField = 'J_NATDEFAUT'
        DataSource = SJal
      end
    end
  end
  object DBNav: TDBNavigator
    Left = 262
    Top = 275
    Width = 80
    Height = 18
    DataSource = SJal
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object BJALDIV: THBitBtn
    Tag = 100
    Left = 311
    Top = 223
    Width = 28
    Height = 16
    Hint = 'Journal divisionnaire'
    Caption = 'JALDIV'
    Enabled = False
    TabOrder = 3
    Visible = False
    OnClick = BJALDIVClick
  end
  object BJALCPT: THBitBtn
    Tag = 100
    Left = 375
    Top = 227
    Width = 28
    Height = 21
    Hint = 'Journal divisionnaire par compte'
    Caption = 'JALDIVCPTE'
    Enabled = False
    TabOrder = 4
    Visible = False
    OnClick = BJALCPTClick
  end
  object BJALPER: THBitBtn
    Tag = 100
    Left = 346
    Top = 226
    Width = 28
    Height = 16
    Hint = 'Journal p'#233'riodique'
    Caption = 'JALPER'
    Enabled = False
    TabOrder = 5
    Visible = False
    OnClick = BJALPERClick
  end
  object BJALCENT: THBitBtn
    Tag = 100
    Left = 380
    Top = 212
    Width = 28
    Height = 16
    Hint = 'Journal centralisateur'
    Caption = 'JALCENT'
    Enabled = False
    TabOrder = 6
    Visible = False
    OnClick = BJALCENTClick
  end
  object BZecrimvt: THBitBtn
    Tag = 100
    Left = 348
    Top = 208
    Width = 27
    Height = 18
    Hint = 'D'#233'tail des mouvements'
    Caption = 'Zmvt'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Visible = False
    OnClick = BZecrimvtClick
    Margin = 1
  end
  object BCumul: THBitBtn
    Tag = 100
    Left = 316
    Top = 204
    Width = 23
    Height = 18
    Hint = 'Cumuls mensuels'
    Caption = 'CumP'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
    Visible = False
    OnClick = BCumulClick
    NumGlyphs = 2
  end
  object SJal: TDataSource
    DataSet = QJal
    OnDataChange = SJalDataChange
    Left = 418
    Top = 32
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;Journaux;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;'
      
        '1;Journaux;Confirmez-vous la suppression de l'#39'enregistrement?;Q;' +
        'YNC;N;C;'
      '2;Journaux;Vous devez renseigner un code.;W;O;O;O;'
      '3;Journaux;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;Journaux;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez l' +
        'e modifier.;W;O;O;O;'
      
        '5;Journaux;Votre saisie est incorrecte : un caract'#232're n'#39'est pas ' +
        'conforme dans la cha'#238'ne.;W;O;O;O;'
      '6;Journaux;Vous devez renseigner un compte existant.;W;O;O;O;'
      
        '7;Journaux;Votre saisie est incorrecte : un caract'#232're n'#39'est pas ' +
        'conforme dans la cha'#238'ne.;W;O;O;O;'
      
        '8;Journaux;Vous ne pouvez pas d'#233'clarer un compte automatique et ' +
        'interdit '#224' la fois;W;O;O;O;'
      '9;Journaux;Vous devez saisir un compte de contrepartie.;W;O;O;O;'
      
        '10;Journaux;Vous devez renseigner un type de contrepartie.;W;O;O' +
        ';O;'
      
        '11;Journaux;Suppression impossible : ce journal est mouvement'#233' p' +
        'ar des '#233'critures analytiques.;W;O;O;O;'
      
        '12;Journaux;Suppression impossible : ce journal est mouvement'#233' p' +
        'ar des '#233'critures comptables.;W;O;O;O;'
      
        '13;Journaux;Suppression impossible : ce journal est mouvement'#233' p' +
        'ar des '#233'critures d'#39'immobilisation.;W;O;O;O;'
      
        '14;Journaux;Suppression impossible : ce journal est un mod'#232'le.;W' +
        ';O;O;O;'
      
        '15;Journaux;Suppression impossible : ce journal est un journal d' +
        #39'ouverture ou de fermeture.;W;O;O;O;'
      
        '16;Journaux;Suppression impossible : ce journal est mouvement'#233' p' +
        'ar des '#233'critures comptables issues de la gestion commerciale.;W;' +
        'O;O;O;'
      'Erreur'
      'L'#39'enregistrement est inaccessible'
      
        '19;Journaux;Ce journal est de nature analytique. Vous devez rens' +
        'eigner un axe analytique.;W;O;O;O;'
      
        '20;Journaux;Vous ne pouvez pas interdire le compte de contrepart' +
        'ie.;W;O;O;O;'
      
        '21;Journaux;Le compte saisi n'#39'est pas valide pour cette r'#233'gulari' +
        'sation.;W;O;O;O;'
      
        '22;Journaux;Ce compte est pr'#233'sent dans les comptes automatiques.' +
        ';W;O;O;O;'
      
        '23;Journaux;Ce compte est pr'#233'sent dans les comptes interdits.;W;' +
        'O;O;O;'
      
        '24;Journaux;Ce journal a un compte de contrepartie en devise. Vo' +
        'us devez renseigner la devise.;W;O;O;O;'
      
        '25;Journaux;Vous devez renseigner un compte de contrepartie exis' +
        'tant.;W;O;O;O;'
      '26;Journaux;Vous devez renseigner un compteur;W;O;O;O;'
      
        '27;Journaux;Vous ne pouvez pas cr'#233'er de jounal en "mode libre";W' +
        ';O;O;O;'
      ''
      '')
    Left = 355
    Top = 98
  end
  object Q1: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 307
    Top = 107
  end
  object Q2: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    MAJAuto = False
    Distinct = False
    Left = 360
    Top = 28
  end
  object PopZ: TPopupMenu
    Left = 394
    Top = 96
  end
  object QJal: TQuery
    AfterPost = QJalAfterPost
    AfterDelete = QJalAfterDelete
    DatabaseName = 'SOC'
    RequestLive = True
    SQL.Strings = (
      'Select * From JOURNAL')
    UpdateMode = upWhereChanged
    Left = 224
    Top = 188
    object QJalJ_JOURNAL: TStringField
      FieldName = 'J_JOURNAL'
      Size = 3
    end
    object QJalJ_LIBELLE: TStringField
      FieldName = 'J_LIBELLE'
      Size = 35
    end
    object QJalJ_ABREGE: TStringField
      FieldName = 'J_ABREGE'
      Size = 17
    end
    object QJalJ_NATUREJAL: TStringField
      FieldName = 'J_NATUREJAL'
      Size = 3
    end
    object QJalJ_DATECREATION: TDateTimeField
      FieldName = 'J_DATECREATION'
      DisplayFormat = 'dd mmm yyyy'
    end
    object QJalJ_DATEOUVERTURE: TDateTimeField
      FieldName = 'J_DATEOUVERTURE'
      DisplayFormat = 'dd mmm yyyy'
    end
    object QJalJ_DATEFERMETURE: TDateTimeField
      FieldName = 'J_DATEFERMETURE'
      DisplayFormat = 'dd mmm yyyy'
    end
    object QJalJ_FERME: TStringField
      FieldName = 'J_FERME'
      Size = 1
    end
    object QJalJ_DATEDERNMVT: TDateTimeField
      FieldName = 'J_DATEDERNMVT'
      DisplayFormat = 'dd mmm yyyy'
    end
    object QJalJ_DEBITDERNMVT: TFloatField
      FieldName = 'J_DEBITDERNMVT'
    end
    object QJalJ_CREDITDERNMVT: TFloatField
      FieldName = 'J_CREDITDERNMVT'
    end
    object QJalJ_NUMDERNMVT: TIntegerField
      FieldName = 'J_NUMDERNMVT'
    end
    object QJalJ_BLOCNOTE: TMemoField
      FieldName = 'J_BLOCNOTE'
      BlobType = ftMemo
      Size = 1
    end
    object QJalJ_COMPTEURNORMAL: TStringField
      FieldName = 'J_COMPTEURNORMAL'
      Size = 3
    end
    object QJalJ_COMPTEURSIMUL: TStringField
      FieldName = 'J_COMPTEURSIMUL'
      Size = 3
    end
    object QJalJ_MULTIDEVISE: TStringField
      FieldName = 'J_MULTIDEVISE'
      Size = 1
    end
    object QJalJ_TYPECONTREPARTIE: TStringField
      FieldName = 'J_TYPECONTREPARTIE'
      Size = 3
    end
    object QJalJ_CONTREPARTIE: TStringField
      FieldName = 'J_CONTREPARTIE'
      Size = 17
    end
    object QJalJ_COMPTEINTERDIT: TStringField
      FieldName = 'J_COMPTEINTERDIT'
      Size = 200
    end
    object QJalJ_CENTRALISABLE: TStringField
      FieldName = 'J_CENTRALISABLE'
      Size = 1
    end
    object QJalJ_UTILISATEUR: TStringField
      FieldName = 'J_UTILISATEUR'
      Size = 3
    end
    object QJalJ_TOTALDEBIT: TFloatField
      FieldName = 'J_TOTALDEBIT'
    end
    object QJalJ_TOTALCREDIT: TFloatField
      FieldName = 'J_TOTALCREDIT'
    end
    object QJalJ_CPTEREGULDEBIT1: TStringField
      FieldName = 'J_CPTEREGULDEBIT1'
      Size = 17
    end
    object QJalJ_CPTEREGULDEBIT2: TStringField
      FieldName = 'J_CPTEREGULDEBIT2'
      Size = 17
    end
    object QJalJ_CPTEREGULDEBIT3: TStringField
      FieldName = 'J_CPTEREGULDEBIT3'
      Size = 17
    end
    object QJalJ_CPTEREGULCREDIT1: TStringField
      FieldName = 'J_CPTEREGULCREDIT1'
      Size = 17
    end
    object QJalJ_CPTEREGULCREDIT2: TStringField
      FieldName = 'J_CPTEREGULCREDIT2'
      Size = 17
    end
    object QJalJ_CPTEREGULCREDIT3: TStringField
      FieldName = 'J_CPTEREGULCREDIT3'
      Size = 17
    end
    object QJalJ_SOCIETE: TStringField
      FieldName = 'J_SOCIETE'
      Size = 3
    end
    object QJalJ_COMPTEAUTOMAT: TStringField
      FieldName = 'J_COMPTEAUTOMAT'
      Size = 200
    end
    object QJalJ_TOTDEBP: TFloatField
      FieldName = 'J_TOTDEBP'
    end
    object QJalJ_TOTCREP: TFloatField
      FieldName = 'J_TOTCREP'
    end
    object QJalJ_TOTDEBE: TFloatField
      FieldName = 'J_TOTDEBE'
    end
    object QJalJ_TOTCREE: TFloatField
      FieldName = 'J_TOTCREE'
    end
    object QJalJ_TOTDEBS: TFloatField
      FieldName = 'J_TOTDEBS'
    end
    object QJalJ_TOTCRES: TFloatField
      FieldName = 'J_TOTCRES'
    end
    object QJalJ_AXE: TStringField
      FieldName = 'J_AXE'
      Size = 3
    end
    object QJalJ_VALIDEEN: TStringField
      FieldName = 'J_VALIDEEN'
      Size = 24
    end
    object QJalJ_VALIDEEN1: TStringField
      FieldName = 'J_VALIDEEN1'
      Size = 24
    end
    object QJalJ_DATEMODIF: TDateTimeField
      FieldName = 'J_DATEMODIF'
      DisplayFormat = 'dd mmm yyyy'
    end
    object QJalJ_EFFET: TStringField
      FieldName = 'J_EFFET'
      Size = 1
    end
    object QJalJ_MODESAISIE: TStringField
      FieldName = 'J_MODESAISIE'
      Origin = 'JOURNAL.J_MODESAISIE'
      Size = 3
    end
    object QJalJ_CHOIXDATE: TStringField
      FieldName = 'J_CHOIXDATE'
      Origin = 'SOC.JOURNAL.J_CHOIXDATE'
      FixedChar = True
      Size = 17
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 271
    Top = 188
  end
end

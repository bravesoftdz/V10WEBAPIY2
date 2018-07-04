object FDrupture: TFDrupture
  Left = 203
  Top = 223
  Width = 455
  Height = 323
  ActiveControl = FListe
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Ruptures sur '
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PAuto: TPanel
    Left = 0
    Top = 0
    Width = 447
    Height = 261
    Align = alClient
    BevelInner = bvLowered
    TabOrder = 3
    Visible = False
    object GBAuto: TGroupBox
      Left = 170
      Top = 49
      Width = 270
      Height = 203
      Caption = ' Libell'#233' des niveaux '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      object TEClasse1: THLabel
        Left = 12
        Top = 25
        Width = 74
        Height = 13
        Caption = 'Libell'#233' niveau 1'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEClasse2: THLabel
        Left = 12
        Top = 55
        Width = 74
        Height = 13
        Caption = 'Libell'#233' niveau 2'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEClasse3: THLabel
        Left = 12
        Top = 84
        Width = 74
        Height = 13
        Caption = 'Libell'#233' niveau 3'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEClasse4: THLabel
        Left = 12
        Top = 114
        Width = 74
        Height = 13
        Caption = 'Libell'#233' niveau 4'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEClasse5: THLabel
        Left = 12
        Top = 143
        Width = 74
        Height = 13
        Caption = 'Libell'#233' niveau 5'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEClasse6: THLabel
        Left = 12
        Top = 173
        Width = 74
        Height = 13
        Caption = 'Libell'#233' niveau 6'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object EClasse1: TEdit
        Left = 97
        Top = 21
        Width = 160
        Height = 21
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object EClasse2: TEdit
        Left = 97
        Top = 51
        Width = 160
        Height = 21
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object EClasse3: TEdit
        Left = 97
        Top = 80
        Width = 160
        Height = 21
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object EClasse4: TEdit
        Left = 97
        Top = 110
        Width = 160
        Height = 21
        Ctl3D = True
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
      object EClasse5: TEdit
        Left = 97
        Top = 139
        Width = 160
        Height = 21
        Ctl3D = True
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
      object EClasse6: TEdit
        Left = 97
        Top = 169
        Width = 160
        Height = 21
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
      end
    end
    object GbAuto1: TGroupBox
      Left = 10
      Top = 49
      Width = 147
      Height = 203
      Caption = ' G'#233'n'#233'rer les niveaux '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object Classe1: TCheckBox
        Left = 31
        Top = 25
        Width = 65
        Height = 18
        Alignment = taLeftJustify
        Caption = 'Niveau 1'
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object Classe3: TCheckBox
        Left = 31
        Top = 84
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Niveau 3'
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object Classe5: TCheckBox
        Left = 31
        Top = 143
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Niveau 5'
        Ctl3D = True
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
      object Classe6: TCheckBox
        Left = 31
        Top = 173
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Niveau 6'
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
      end
      object Classe4: TCheckBox
        Left = 31
        Top = 114
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Niveau 4'
        Ctl3D = True
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
      object Classe2: TCheckBox
        Left = 31
        Top = 55
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Niveau 2'
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
    end
    object Panel1: TPanel
      Left = 2
      Top = 2
      Width = 443
      Height = 41
      Align = alTop
      BevelInner = bvLowered
      TabOrder = 0
      object LClasse: TLabel
        Left = 120
        Top = 14
        Width = 134
        Height = 13
        Alignment = taCenter
        Caption = '&Choix du nombre de niveaux'
      end
      object SEAuto: TSpinEdit
        Left = 267
        Top = 10
        Width = 36
        Height = 22
        Ctl3D = True
        EditorEnabled = False
        MaxLength = 1
        MaxValue = 6
        MinValue = 0
        ParentCtl3D = False
        TabOrder = 0
        Value = 0
        OnChange = SEAutoChange
      end
    end
  end
  object FListe: THDBGrid
    Left = 0
    Top = 0
    Width = 447
    Height = 261
    Align = alClient
    DataSource = SRupture
    Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyDown = FListeKeyDown
    OnRowEnter = FListeRowEnter
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
    Columns = <
      item
        Expanded = False
        FieldName = 'RU_CLASSE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ReadOnly = True
        Title.Alignment = taCenter
        Title.Caption = 'Classe'
        Width = 138
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'RU_LIBELLECLASSE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 279
        Visible = True
      end>
  end
  object DBNav: TDBNavigator
    Left = 144
    Top = 208
    Width = 80
    Height = 18
    DataSource = SRupture
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object HPB: TPanel
    Left = 0
    Top = 261
    Width = 447
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
    TabOrder = 1
    object FAutoSave: TCheckBox
      Left = 112
      Top = 8
      Width = 25
      Height = 17
      Caption = 'Auto'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 0
      Visible = False
    end
    object BAnnuler: THBitBtn
      Left = 120
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Annuler la derni'#232're action'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
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
      TabOrder = 1
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrev: THBitBtn
      Left = 32
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Pr'#233'c'#233'dent'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BPrevClick
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
      TabOrder = 4
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
    object BInsert: THBitBtn
      Left = 152
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Nouveau'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = BInsertClick
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object BDelete: THBitBtn
      Left = 184
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Supprimer'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = BDeleteClick
      Margin = 0
      GlobalIndexImage = 'Z0005_S16G1'
    end
    object BAutomate: THBitBtn
      Left = 248
      Top = 4
      Width = 28
      Height = 27
      Hint = 'G'#233'n'#233'ration automatique des ruptures'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      OnClick = BAutomateClick
      GlobalIndexImage = 'Z0775_S16G1'
    end
    object BRuptanal: THBitBtn
      Left = 216
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Enchainements des ruptures analytiques'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = BRuptanalClick
      GlobalIndexImage = 'Z0566_S16G1'
    end
    object Panel2: TPanel
      Left = 319
      Top = 2
      Width = 126
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 10
      object BAide: THBitBtn
        Left = 95
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
        Left = 65
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
        Left = 34
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
      object BImprimer: THBitBtn
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
    object BZoomSp: THBitBtn
      Left = 280
      Top = 4
      Width = 28
      Height = 27
      Hint = 
        'G'#233'n'#233'ration automatique des ruptures '#224' partir des sous plans anal' +
        'ytiques'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 11
      OnClick = BZoomSpClick
      GlobalIndexImage = 'Z0192_S16G1'
    end
  end
  object SRupture: TDataSource
    DataSet = TRupture
    OnDataChange = SRuptureDataChange
    OnUpdateData = SRuptureUpdateData
    Left = 113
    Top = 116
  end
  object TRupture: THTable
    MarshalOptions = moMarshalModifiedOnly
    BeforePost = TRuptureBeforePost
    OnNewRecord = TRuptureNewRecord
    OnPostError = TRupturePostError
    EnableBCD = False
    IndexName = 'RU_CLE1'
    TableName = 'RUPTURE'
    dataBaseName = 'SOC'
    RequestLive = True
    Left = 17
    Top = 116
    object TRuptureRU_NATURERUPT: TStringField
      FieldName = 'RU_NATURERUPT'
      Size = 3
    end
    object TRuptureRU_CLASSE: TStringField
      FieldName = 'RU_CLASSE'
      EditMask = '>a;0;  '
      Size = 17
    end
    object TRuptureRU_LIBELLECLASSE: TStringField
      FieldName = 'RU_LIBELLECLASSE'
      Size = 35
    end
    object TRuptureRU_SOCIETE: TStringField
      FieldName = 'RU_SOCIETE'
      Size = 3
    end
    object TRuptureRU_PLANRUPT: TStringField
      FieldName = 'RU_PLANRUPT'
      Size = 3
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Plans de ruptures des '#233'ditions;D'#233'sirez-vous enregistrer les mo' +
        'difications?;Q;YNC;Y;C;'
      
        '1;Plans de ruptures des '#233'ditions;Confirmez-vous la suppression d' +
        'e l'#39'enregistrement?;Q;YNC;N;C;'
      
        '2;Plans de ruptures des '#233'ditions;Vous devez renseigner un code.;' +
        'W;O;O;O;'
      
        '3;Plans de ruptures des '#233'ditions;Vous devez renseigner un libell' +
        #233'.;W;O;O;O;'
      
        '4;Plans de ruptures des '#233'ditions;Le code choisi existe d'#233'j'#224'. Vou' +
        's devez le modifier.;W;O;O;O;'
      
        '5;Plans de ruptures des '#233'ditions;Confirmez-vous la destruction d' +
        'es anciennes et la cr'#233'ation des nouvelles classes de rupture?;Q;' +
        'YNC;Y;C;'
      
        '6;Plans de ruptures des '#233'ditions;D'#233'sirez-vous enregistrer les mo' +
        'difications;Q;YNC;Y;C;'
      
        '7;Plans de ruptures des '#233'ditions;Aucun encha'#238'nement existe pour ' +
        'ce plan. Vous devez le cr'#233'er.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      'comptes g'#233'n'#233'raux'
      'comptes auxiliaires'
      'sections axe 1'
      'sections axe 2'
      'sections axe 3'
      'sections axe 4'
      'sections axe 5'
      
        '16;Plans de ruptures des '#233'ditions;Les ruptures n'#39'ont pas '#233't'#233' g'#233'n' +
        #233'r'#233'es. La modification de l'#39'encha'#238'nement n'#39'est pas valid'#233'e.;W;O;' +
        'O;O;'
      'comptes budg'#233'taires')
    Left = 68
    Top = 168
  end
  object QRupture: TQuery
    DatabaseName = 'SOC'
    Left = 24
    Top = 164
  end
  object HMTrad: THSystemMenu
    ActiveResize = False
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 68
    Top = 116
  end
end

object FRepartcle: TFRepartcle
  Left = 246
  Top = 135
  HelpContext = 1455100
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  ActiveControl = PAxe
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Cl'#233' de r'#233'partition analytique : '
  ClientHeight = 229
  ClientWidth = 577
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pappli: TPanel
    Left = 0
    Top = 0
    Width = 375
    Height = 190
    Align = alClient
    Caption = ' '
    Ctl3D = True
    ParentCtl3D = False
    TabOrder = 0
    object TRE_LIBELLE: THLabel
      Left = 7
      Top = 71
      Width = 30
      Height = 13
      Caption = '&Libell'#233
      FocusControl = RE_LIBELLE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TRE_MODEREPARTITION: TLabel
      Left = 7
      Top = 97
      Width = 50
      Height = 26
      Caption = '&Modes de r'#233'partition'
      FocusControl = RE_MODEREPARTITION
      WordWrap = True
    end
    object TRE_QUALIFQTE: TLabel
      Left = 7
      Top = 131
      Width = 44
      Height = 26
      Caption = '&Qualifiant quantit'#233's'
      FocusControl = RE_QUALIFQTE
      WordWrap = True
    end
    object TRE_CLE: TLabel
      Left = 8
      Top = 43
      Width = 25
      Height = 13
      Caption = '&Code'
      FocusControl = RE_CLE
    end
    object TRE_ABREGE: TLabel
      Left = 168
      Top = 43
      Width = 34
      Height = 13
      Caption = '&Abr'#233'g'#233
      FocusControl = RE_ABREGE
    end
    object TRE_COMPTES: TLabel
      Left = 7
      Top = 161
      Width = 44
      Height = 26
      Caption = '&Filtre g'#233'n'#233'raux'
      FocusControl = RE_COMPTES
      WordWrap = True
    end
    object RE_CLE: TDBEdit
      Left = 67
      Top = 40
      Width = 56
      Height = 21
      CharCase = ecUpperCase
      Ctl3D = True
      DataField = 'RE_CLE'
      DataSource = SRepartCle
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
    end
    object RE_LIBELLE: TDBEdit
      Left = 67
      Top = 68
      Width = 294
      Height = 21
      Ctl3D = True
      DataField = 'RE_LIBELLE'
      DataSource = SRepartCle
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
    end
    object RE_ABREGE: TDBEdit
      Left = 215
      Top = 40
      Width = 146
      Height = 21
      Ctl3D = True
      DataField = 'RE_ABREGE'
      DataSource = SRepartCle
      ParentCtl3D = False
      TabOrder = 1
    end
    object RE_MODEREPARTITION: THDBValComboBox
      Left = 67
      Top = 100
      Width = 294
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 3
      TagDispatch = 0
      DataType = 'TTMODEREPARTANAL'
      DataField = 'RE_MODEREPARTITION'
      DataSource = SRepartCle
    end
    object RE_QUALIFQTE: THDBValComboBox
      Left = 67
      Top = 132
      Width = 294
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 4
      TagDispatch = 0
      DataType = 'TTQUALUNITMESURE'
      DataField = 'RE_QUALIFQTE'
      DataSource = SRepartCle
    end
    object PAxe: TTabControl
      Left = 1
      Top = 1
      Width = 373
      Height = 29
      Align = alTop
      TabOrder = 5
      Tabs.Strings = (
        'Axe n'#176'1'
        'Axe n'#176'2'
        'Axe n'#176'3'
        'Axe n'#176'4'
        'Axe n'#176'5')
      TabIndex = 0
      OnChange = PAxeChange
    end
    object RE_COMPTES: TDBEdit
      Left = 67
      Top = 164
      Width = 294
      Height = 21
      Ctl3D = True
      DataField = 'RE_COMPTES'
      DataSource = SRepartCle
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 6
    end
  end
  object FListe: THDBGrid
    Left = 375
    Top = 0
    Width = 202
    Height = 190
    Align = alRight
    Ctl3D = True
    DataSource = SRepartCle
    Options = [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgCancelOnExit]
    ParentCtl3D = False
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyDown = FListeKeyDown
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
    Columns = <
      item
        Expanded = False
        FieldName = 'RE_CLE'
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 33
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'RE_LIBELLE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 135
        Visible = True
      end>
  end
  object DBNav: TDBNavigator
    Left = 428
    Top = 105
    Width = 80
    Height = 18
    DataSource = SRepartCle
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    Hints.Strings = (
      'Premier'
      'Pr'#233'c'#233'dent'
      'Suivant'
      'Dernier')
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object Dock: TDock97
    Left = 0
    Top = 190
    Width = 577
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 573
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 573
      DockPos = 0
      TabOrder = 0
      object BInsert: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Caption = 'Nouveau'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsertClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 68
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Supprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BFirst: TToolbarButton97
        Left = 380
        Top = 3
        Width = 48
        Height = 27
        Hint = 'Premier'
        DisplayMode = dmGlyphOnly
        Caption = 'Premier'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFirstClick
        GlobalIndexImage = 'Z0301_S16G1'
      end
      object BPrev: TToolbarButton97
        Left = 428
        Top = 3
        Width = 48
        Height = 27
        Hint = 'Pr'#233'c'#233'dent'
        DisplayMode = dmGlyphOnly
        Caption = 'Pr'#233'c'#233'dent'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPrevClick
        GlobalIndexImage = 'Z0017_S16G1'
      end
      object BNext: TToolbarButton97
        Left = 476
        Top = 3
        Width = 48
        Height = 27
        Hint = 'Suivant'
        DisplayMode = dmGlyphOnly
        Caption = 'Suivant'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNextClick
        GlobalIndexImage = 'Z0031_S16G1'
      end
      object BLast: TToolbarButton97
        Left = 524
        Top = 3
        Width = 48
        Height = 27
        Hint = 'Dernier'
        DisplayMode = dmGlyphOnly
        Caption = 'Dernier'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BLastClick
        GlobalIndexImage = 'Z0264_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 344
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 280
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
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
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Annuler la derni'#232're action'
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 312
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
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
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
      end
      object BImprimer: TToolbarButton97
        Left = 248
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BVentil: TToolbarButton97
        Left = 172
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Sections r'#233'ceptrices'
        DisplayMode = dmGlyphOnly
        Caption = 'Sections'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BVentilClick
        GlobalIndexImage = 'Z0133_S16G1'
        IsControl = True
      end
    end
  end
  object FAutoSave: TCheckBox
    Left = 117
    Top = 36
    Width = 36
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
    TabOrder = 4
    Visible = False
  end
  object SRepartCle: TDataSource
    DataSet = TRepartCle
    OnStateChange = SRepartCleStateChange
    OnDataChange = SRepartCleDataChange
    OnUpdateData = SRepartCleUpdateData
    Left = 402
    Top = 48
  end
  object TRepartCle: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TRepartCleAfterPost
    AfterDelete = TRepartCleAfterDelete
    EnableBCD = False
    IndexName = 'RE_CLE1'
    TableName = 'CLEREPAR'
    dataBaseName = 'SOC'
    RequestLive = True
    Left = 462
    Top = 48
    object TRepartCleRE_AXE: TStringField
      FieldName = 'RE_AXE'
      Size = 3
    end
    object TRepartCleRE_CLE: TStringField
      FieldName = 'RE_CLE'
      Size = 3
    end
    object TRepartCleRE_LIBELLE: TStringField
      FieldName = 'RE_LIBELLE'
      Size = 35
    end
    object TRepartCleRE_ABREGE: TStringField
      FieldName = 'RE_ABREGE'
      Size = 17
    end
    object TRepartCleRE_MODEREPARTITION: TStringField
      FieldName = 'RE_MODEREPARTITION'
      Size = 3
    end
    object TRepartCleRE_QUALIFQTE: TStringField
      FieldName = 'RE_QUALIFQTE'
      Size = 3
    end
    object TRepartCleRE_COMPTES: TStringField
      FieldName = 'RE_COMPTES'
      Size = 200
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Cl'#233's de r'#233'partition analytique;Voulez-vous enregistrer les mod' +
        'ifications?;Q;YNC;Y;C;'
      
        '1;Cl'#233's de r'#233'partition analytique;Confirmez-vous la suppression d' +
        'e l'#39'enregistrement?;Q;YNC;N;C;'
      
        '2;Cl'#233's de r'#233'partition analytique;Vous devez renseigner un code.;' +
        'W;O;O;O;'
      
        '3;Cl'#233's de r'#233'partition analytique;Vous devez renseigner un libell' +
        #233'.;W;O;O;O;'
      
        '4;Cl'#233's de r'#233'partition analytique;Le code que vous avez renseign'#233 +
        ' existe d'#233'j'#224' sur cet axe. Vous devez le modifier.;W;O;O;O;'
      
        '5;Cl'#233's de r'#233'partition analytique;La somme des taux de r'#233'partitio' +
        'n doit '#234'tre '#233'gale '#224' 100%;W;O;O;O;'
      
        '6;Cl'#233's de r'#233'partition analytique;Vous devez renseigner un exerci' +
        'ce.;W;O;O;O;'
      
        '7;Cl'#233's de r'#233'partition analytique;Vous devez renseigner un axe an' +
        'alytique.;W;O;O;O;'
      
        '8;Cl'#233's de r'#233'partition analytique;Vous devez renseigner un mode d' +
        'e r'#233'partition.;W;O;O;O;'
      
        '9;Cl'#233's de r'#233'partition analytique;Vous devez renseigner au moins ' +
        'une section analytique;W;O;O;O;'
      
        '10;Cl'#233's de r'#233'partition analytique;Les sections doivent '#234'tre rens' +
        'eign'#233'es dans l'#39'ordre;W;O;O;O;'
      'L'#39'enregistrement est inaccessible')
    Left = 514
    Top = 48
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 520
    Top = 104
  end
end

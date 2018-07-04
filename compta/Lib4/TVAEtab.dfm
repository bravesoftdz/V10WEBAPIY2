object FTvaEtab: TFTvaEtab
  Left = 338
  Top = 177
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'R'#233'gime par '#233'tablissement:'
  ClientHeight = 323
  ClientWidth = 557
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
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
  object FListe2: THDBGrid
    Left = 379
    Top = 0
    Width = 178
    Height = 287
    Align = alRight
    Ctl3D = True
    DataSource = STEtab
    Options = [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection]
    ParentCtl3D = False
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clBlack
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnCellClick = FListe2CellClick
    OnKeyDown = FListe2KeyDown
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
    Columns = <
      item
        Expanded = False
        FieldName = 'ET_ETABLISSEMENT'
        Title.Caption = 'Code'
        Width = 65
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'ET_ABREGE'
        Title.Caption = 'Libelle'
        Visible = True
      end>
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 379
    Height = 287
    Align = alClient
    Caption = ' '
    TabOrder = 1
    object HB1: TPanel
      Left = 1
      Top = 1
      Width = 377
      Height = 56
      Align = alTop
      Caption = ' '
      TabOrder = 0
      object TCC_CODE: THLabel
        Left = 8
        Top = 32
        Width = 25
        Height = 13
        Caption = '&Code'
        FocusControl = ET_ETABLISSEMENT
      end
      object TCC_LIBELLE: THLabel
        Left = 93
        Top = 32
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = ET_LIBELLE
      end
      object TTVC_CODETAUX: THLabel
        Left = 8
        Top = 7
        Width = 74
        Height = 13
        Caption = 'Code TVA/TPF'
      end
      object ET_ETABLISSEMENT: TDBEdit
        Left = 40
        Top = 28
        Width = 37
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'ET_ETABLISSEMENT'
        DataSource = STEtab
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object ET_LIBELLE: TDBEdit
        Left = 135
        Top = 28
        Width = 231
        Height = 21
        Ctl3D = True
        DataField = 'ET_LIBELLE'
        DataSource = STEtab
        ParentCtl3D = False
        TabOrder = 1
      end
      object Taux: TEdit
        Left = 135
        Top = 4
        Width = 231
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ReadOnly = True
        TabOrder = 2
      end
    end
    object DBNav: TDBNavigator
      Left = 148
      Top = 92
      Width = 80
      Height = 18
      DataSource = STEtab
      VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
      ConfirmDelete = False
      TabOrder = 1
      Visible = False
    end
    object Cache: THCpteEdit
      Left = 118
      Top = 91
      Width = 17
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      Text = '!!!'
      Visible = False
      ZoomTable = tzGNonCollectif
      Vide = True
      Bourre = True
      okLocate = True
      SynJoker = False
    end
    object FListe: THDBGrid
      Left = 1
      Top = 57
      Width = 377
      Height = 229
      Align = alClient
      DataSource = STauxCompl
      Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgCancelOnExit]
      TabOrder = 3
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clBlack
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnDblClick = FListeDblClick
      OnEnter = FListeEnter
      OnKeyDown = FListeKeyDown
      OnKeyPress = FListeKeyPress
      Row = 1
      MultiSelection = False
      SortEnabled = False
      MyDefaultRowHeight = 0
      Columns = <
        item
          Expanded = False
          FieldName = 'TVC_TVAOUTPF'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'TVC_CODETAUX'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'TVC_REGIME'
          Title.Caption = 'R'#233'gime'
          Width = 81
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TVC_ETABLISSEMENT'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'TVC_CPTEACH'
          Title.Caption = 'Compte TVA '#224' l'#39'achat'
          Width = 128
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TVC_CPTEVTE'
          Title.Caption = 'Compte TVA '#224' la vente'
          Width = 135
          Visible = True
        end>
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 287
    Width = 557
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 557
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 557
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 339
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
      object BValider: TToolbarButton97
        Left = 275
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
      object BImprimer: TToolbarButton97
        Left = 243
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
      object BFirst: TToolbarButton97
        Left = 384
        Top = 3
        Width = 38
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
        Left = 423
        Top = 3
        Width = 38
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
        Left = 462
        Top = 3
        Width = 38
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
        Left = 501
        Top = 3
        Width = 38
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
      object BInsert: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Caption = 'Nouveau'
        Flat = False
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
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BFerme: TToolbarButton97
        Left = 307
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
    end
  end
  object FAutoSave: TCheckBox
    Left = 160
    Top = 120
    Width = 33
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
    TabOrder = 3
    Visible = False
  end
  object TEtab: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TEtabAfterPost
    EnableBCD = False
    IndexName = 'ET_CLE1'
    TableName = 'ETABLISS'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 476
    Top = 44
    object TEtabET_ETABLISSEMENT: TStringField
      FieldName = 'ET_ETABLISSEMENT'
      Size = 3
    end
    object TEtabET_ABREGE: TStringField
      FieldName = 'ET_ABREGE'
      Size = 17
    end
    object TEtabET_LIBELLE: TStringField
      FieldName = 'ET_LIBELLE'
      Size = 35
    end
  end
  object STEtab: TDataSource
    DataSet = TEtab
    OnStateChange = STEtabStateChange
    OnDataChange = STEtabDataChange
    Left = 520
    Top = 44
  end
  object TTauxCompl: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TTauxComplAfterPost
    AfterDelete = TTauxComplAfterDelete
    OnNewRecord = TTauxComplNewRecord
    EnableBCD = False
    IndexName = 'TVC_CLE1'
    TableName = 'TXCPTTVACOMPL'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 224
    Top = 84
    object TTauxComplTVC_TVAOUTPF: TStringField
      FieldName = 'TVC_TVAOUTPF'
      Size = 3
    end
    object TTauxComplTVC_CODETAUX: TStringField
      FieldName = 'TVC_CODETAUX'
      Size = 3
    end
    object TTauxComplTVC_REGIME: TStringField
      FieldName = 'TVC_REGIME'
      OnChange = TTauxComplTVC_REGIMEChange
      Size = 3
    end
    object TTauxComplTVC_ETABLISSEMENT: TStringField
      FieldName = 'TVC_ETABLISSEMENT'
      Size = 3
    end
    object TTauxComplTVC_CPTEACH: TStringField
      FieldName = 'TVC_CPTEACH'
      Size = 17
    end
    object TTauxComplTVC_CPTEVTE: TStringField
      FieldName = 'TVC_CPTEVTE'
      Size = 17
    end
  end
  object STauxCompl: TDataSource
    DataSet = TTauxCompl
    OnStateChange = STauxComplStateChange
    OnDataChange = STauxComplDataChange
    OnUpdateData = STauxComplUpdateData
    Left = 320
    Top = 84
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C' +
        ';'
      
        '1;?caption?;Confirmez-vous la suppression de l'#39'enregistrement?;Q' +
        ';YNC;N;C;'
      '2;?caption?;Vous devez renseigner un code.;W;O;O;O;'
      '3;?caption?;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;?caption?;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez ' +
        'le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '6;?caption?;Le compte '#224' l'#39'achat d'#233'bit que vous avez saisi n'#39'exis' +
        'te pas, est incompatible ou incomplet;W;O;O;O;'
      
        '7;?caption?;Le compte '#224' la vente d'#233'bit que vous avez saisi n'#39'exi' +
        'ste pas, est incompatible ou incomplet.;W;O;O;O;'
      '8;?caption?;Le regime est obligatoire;W;O;O;O;'
      
        '9;?caption?;Vous devez renseigner au moins un compte d'#39'achat ou ' +
        'de vente;W;O;O;O;'
      '10;?caption?;Vous devez renseigner le compte '#224' l'#39'achat.;W;O;O;O;'
      
        '11;?caption?;Vous devez renseigner le compte '#224' la vente.;W;O;O;O' +
        ';'
      '12;?caption?;L'#39'enregistrement ne sera pas sauvegard'#233'.;W;O;O;O;'
      
        '13;?caption?;Le code est r'#233'f'#233'renc'#233' par un compte g'#233'n'#233'ral.;W;O;O;' +
        'O;'
      
        '14;?caption?;Le compte '#224' l'#39'achat encaissement que vous avez sais' +
        'i n'#39'existe pas, est incompatibe ou incomplet;W;O;O;O;'
      
        '15;?caption?;Le compte '#224' la vente encaissement que vous avez sai' +
        'si n'#39'existe pas, est incompatible ou incomplet;W;O;O;O;'
      '16;?caption?;Cet indice de code TVA est d'#233'j'#224' attribu'#233';W;O;O;O;'
      '17;?caption?;Ce code regime existe d'#233'j'#224';W;O;O;O;')
    Left = 36
    Top = 88
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    LockedCtrls.Strings = (
      'FListe')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 512
    Top = 104
  end
  object HT: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Compte TPF '#224' l'#39'achat'
      'Compte TPF '#224' la vente')
    Left = 68
    Top = 96
  end
end

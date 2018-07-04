object FRupture: TFRupture
  Left = 371
  Top = 171
  Width = 411
  Height = 223
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  ActiveControl = Fliste
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Plans de ruptures des '#233'ditions sur '
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
  object Dock: TDock97
    Left = 0
    Top = 157
    Width = 403
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 403
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 403
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BDefaire: TToolbarButton97
        Left = 120
        Top = 2
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
        OnClick = BDefaireClick
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object BFirst: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
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
        Left = 32
        Top = 2
        Width = 28
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
        Left = 60
        Top = 2
        Width = 28
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
        Left = 88
        Top = 2
        Width = 28
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
        Left = 152
        Top = 2
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
        Left = 184
        Top = 2
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
      object BDplanRup: TToolbarButton97
        Left = 216
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'tail des plans de ruptures'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'tail'
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
        OnClick = BDplanRupClick
        GlobalIndexImage = 'Z0133_S16G1'
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 273
        Top = 2
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
      object BValider: TToolbarButton97
        Left = 305
        Top = 2
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
      object BFerme: TToolbarButton97
        Left = 336
        Top = 2
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
      object BAide: TToolbarButton97
        Left = 368
        Top = 2
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
    end
  end
  object Fliste: THDBGrid
    Left = 0
    Top = 0
    Width = 403
    Height = 157
    Align = alClient
    DataSource = SChoixcod
    Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = FlisteDblClick
    OnKeyDown = FlisteKeyDown
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
    Columns = <
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CC_CODE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ReadOnly = True
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 57
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CC_LIBELLE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 312
        Visible = True
      end>
  end
  object DBNav: TDBNavigator
    Left = 16
    Top = 125
    Width = 80
    Height = 18
    DataSource = SChoixcod
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object FAutoSave: TCheckBox
    Left = 296
    Top = 68
    Width = 29
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
  object SChoixcod: TDataSource
    DataSet = TChoixcod
    OnDataChange = SChoixcodDataChange
    OnUpdateData = SChoixcodUpdateData
    Left = 124
    Top = 60
  end
  object TChoixcod: THTable
    MarshalOptions = moMarshalModifiedOnly
    BeforeDelete = TChoixcodBeforeDelete
    OnNewRecord = TChoixcodNewRecord
    OnPostError = TChoixcodPostError
    EnableBCD = False
    IndexName = 'CC_CLE1'
    TableName = 'CHOIXCOD'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 184
    Top = 60
    object TChoixcodCC_TYPE: TStringField
      FieldName = 'CC_TYPE'
      Size = 3
    end
    object TChoixcodCC_CODE: TStringField
      FieldName = 'CC_CODE'
      EditMask = '>aaa;0; '
      Size = 3
    end
    object TChoixcodCC_LIBELLE: TStringField
      FieldName = 'CC_LIBELLE'
      Size = 35
    end
    object TChoixcodCC_LIBRE: TStringField
      FieldName = 'CC_LIBRE'
      Size = 70
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      
        '1;?caption?;Confirmez-vous la suppression de l'#39'enregistrement ?;' +
        'Q;YNC;N;C;'
      '2;?caption?;Vous devez renseigner un code.;W;O;O;O;'
      '3;?caption?;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;?caption?;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez ' +
        'le modifier.;W;O;O;O;'
      
        '5;?caption?;D'#233'sirez-vous cr'#233'er des encha'#238'nements de ruptures ?;Q' +
        ';YNC;Y;C;'
      'L'#39'enregistrement est inaccessible'
      
        '7;?caption?;L'#39'enregistrement n'#39'a pas pu '#234'tre sauvegard'#233' : le cod' +
        'e est vide ou existe d'#233'j'#224'.;E;O;O;O;'
      'les comptes g'#233'n'#233'raux'
      'les comptes auxiliaires'
      'les sections de l'#39'axe 1'
      'les sections de l'#39'axe 2'
      'les sections de l'#39'axe 3'
      'les sections de l'#39'axe 4'
      'les sections de l'#39'axe 5'
      
        '15;?caption?;Les ruptures n'#39'ont pas '#233't'#233' g'#233'n'#233'r'#233'es. La modificatio' +
        'n de l'#39'encha'#238'nement n'#39'est pas valid'#233'e.;W;O;O;O;'
      
        '16;?caption?;D'#233'sirez-vous cr'#233'er des encha'#238'nements de ruptures ?;' +
        'Q;YNC;Y;C;'
      
        '17;?caption?;D'#233'sirez-vous cr'#233'er des encha'#238'nements de ruptures ?;' +
        'Q;YNC;Y;C;'
      'les comptes budg'#233'taires')
    Left = 236
    Top = 60
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 68
    Top = 60
  end
end

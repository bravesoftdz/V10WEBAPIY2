object FVentilType: TFVentilType
  Left = 238
  Top = 145
  Width = 436
  Height = 317
  HelpContext = 1460000
  BorderIcons = [biSystemMenu]
  Caption = 'Ventilations types'
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
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object FListe: THDBGrid
    Left = 0
    Top = 0
    Width = 428
    Height = 248
    Align = alClient
    DataSource = SChoixCod
    Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
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
        FieldName = 'CC_CODE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ReadOnly = True
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 47
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CC_LIBELLE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 242
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CC_ABREGE'
        Title.Alignment = taCenter
        Title.Caption = 'Abr'#233'g'#233
        Width = 105
        Visible = True
      end>
  end
  object Dock: TDock97
    Left = 0
    Top = 248
    Width = 428
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 424
      Caption = 'Actions'
      ClientAreaHeight = 31
      ClientAreaWidth = 424
      DockPos = 0
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
      object BVentil: TToolbarButton97
        Left = 216
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Param'#233'trage des ventilations'
        DisplayMode = dmGlyphOnly
        Caption = 'Ventilations'
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
      object BImprimer: TToolbarButton97
        Left = 298
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
        Left = 330
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
        Left = 362
        Top = 2
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
      object BAide: TToolbarButton97
        Left = 394
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
  object DBNav: TDBNavigator
    Left = 220
    Top = 108
    Width = 80
    Height = 18
    DataSource = SChoixCod
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object FAutoSave: TCheckBox
    Left = 102
    Top = 80
    Width = 16
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
  object SChoixCod: TDataSource
    DataSet = TChoixCod
    OnDataChange = SChoixCodDataChange
    OnUpdateData = SChoixCodUpdateData
    Left = 129
    Top = 204
  end
  object TChoixCod: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TChoixCodAfterPost
    AfterDelete = TChoixCodAfterDelete
    OnNewRecord = TChoixCodNewRecord
    OnPostError = TChoixCodPostError
    EnableBCD = False
    IndexName = 'CC_CLE1'
    TableName = 'CHOIXCOD'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 71
    Top = 204
    object TChoixCodCC_TYPE: TStringField
      FieldName = 'CC_TYPE'
      Size = 3
    end
    object TChoixCodCC_CODE: TStringField
      FieldName = 'CC_CODE'
      EditMask = '>aaa;0; '
      Size = 3
    end
    object TChoixCodCC_LIBELLE: TStringField
      FieldName = 'CC_LIBELLE'
      Size = 35
    end
    object TChoixCodCC_ABREGE: TStringField
      FieldName = 'CC_ABREGE'
      Size = 17
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Ventilations types;Voulez-vous enregistrer les modifications?;' +
        'Q;YNC;Y;C;'
      
        '1;Ventilations types;Confirmez-vous la suppression de l'#39'enregist' +
        'rement?;Q;YNC;N;C;'
      '2;Ventilations types;Vous devez renseigner un code.;W;O;O;O;'
      '3;Ventilations types;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;Ventilations types;Le code que vous avez saisi existe d'#233'j'#224'. Vo' +
        'us devez le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible')
    Left = 12
    Top = 204
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 188
    Top = 204
  end
end

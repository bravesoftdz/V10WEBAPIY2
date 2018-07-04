object FCodeAFB: TFCodeAFB
  Left = 226
  Top = 172
  Width = 511
  Height = 317
  HelpContext = 1705040
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Codes AFB'
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
  object FListe: THDBGrid
    Left = 0
    Top = 0
    Width = 503
    Height = 248
    Align = alClient
    DataSource = SCodeAFB
    Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clBlack
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
        Alignment = taCenter
        Expanded = False
        FieldName = 'AF_CODEAFB'
        Title.Alignment = taCenter
        Title.Caption = 'Code AFB'
        Width = 55
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AF_LIBELLE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233' AFB'
        Width = 180
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AF_MODEPAIEMENT'
        Title.Alignment = taCenter
        Title.Caption = 'Code M.P.'
        Width = 80
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AF_MPLibelle'
        ReadOnly = False
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233' M.P.'
        Width = 150
        Visible = True
      end>
  end
  object Dock: TDock97
    Left = 0
    Top = 248
    Width = 503
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 503
      Caption = 'Actions'
      ClientAreaHeight = 31
      ClientAreaWidth = 503
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BAnnuler: TToolbarButton97
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
        OnClick = BAnnulerClick
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
      object BImprimer: TToolbarButton97
        Left = 372
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
        Left = 404
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
        Left = 436
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
        Left = 468
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
    DataSource = SCodeAFB
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object FAutoSave: TCheckBox
    Left = 324
    Top = 112
    Width = 45
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
  object SCodeAFB: TDataSource
    DataSet = TCodeAFB
    OnStateChange = SCodeAFBStateChange
    OnDataChange = SCodeAFBDataChange
    OnUpdateData = SCodeAFBUpdateData
    Left = 112
    Top = 128
  end
  object TCodeAFB: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TCodeAFBAfterPost
    AfterDelete = TCodeAFBAfterDelete
    OnNewRecord = TCodeAFBNewRecord
    OnPostError = TCodeAFBPostError
    EnableBCD = False
    IndexName = 'AF_CLE1'
    TableName = 'CODEAFB'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 44
    Top = 128
    object TCodeAFBAF_CODEAFB: TStringField
      FieldName = 'AF_CODEAFB'
      Size = 3
    end
    object TCodeAFBAF_LIBELLE: TStringField
      FieldName = 'AF_LIBELLE'
      Size = 35
    end
    object TCodeAFBAF_MODEPAIEMENT: TStringField
      FieldName = 'AF_MODEPAIEMENT'
      Size = 3
    end
    object TCodeAFBAF_MPLibelle: TStringField
      FieldKind = fkLookup
      FieldName = 'AF_MPLibelle'
      LookupDataSet = TModePaie
      LookupKeyFields = 'MP_MODEPAIE'
      LookupResultField = 'MP_LIBELLE'
      KeyFields = 'AF_MODEPAIEMENT'
      ReadOnly = True
      Size = 35
      Lookup = True
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Codes AFB;Voulez vous enregistrer les modifications?;Q;YNC;Y;C' +
        ';'
      
        '1;Codes AFB;Confirmez-vous la suppression de l'#39'enregistrement?;Q' +
        ';YNC;N;C;'
      '2;Codes AFB;Vous devez renseigner un code;W;O;O;O;'
      '3;Codes AFB;Vous devez renseigner un libell'#233';W;O;O;O;'
      
        '4;Codes AFB;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez ' +
        'le modifier.;W;O;O;O;'
      '5;Codes AFB;Ce mode de paiement n'#39'existe pas;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '7;Codes AFB;L'#39'enregistrement n'#39'a pu '#234'tre sauvegard'#233'. V'#233'rifiez qu' +
        'e le code n'#39'est ni vide, ni d'#233'j'#224' existant.;W;O;O;O;')
    Left = 24
    Top = 40
  end
  object TModePaie: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    TableName = 'MODEPAIE'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 88
    Top = 64
    object TModePaieMP_MODEPAIE: TStringField
      FieldName = 'MP_MODEPAIE'
      Size = 3
    end
    object TModePaieMP_LIBELLE: TStringField
      FieldName = 'MP_LIBELLE'
      Size = 35
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 380
    Top = 4
  end
end

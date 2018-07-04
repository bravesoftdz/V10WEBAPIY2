object FStrucBud: TFStrucBud
  Left = 168
  Top = 155
  Width = 540
  Height = 330
  HelpContext = 7577200
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Cat'#233'gories de budget :'
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object FListe: THDBGrid
    Left = 0
    Top = 0
    Width = 532
    Height = 263
    Align = alClient
    DataSource = Sa
    Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgCancelOnExit]
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
        Width = 55
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CC_LIBELLE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 289
        Visible = True
      end>
  end
  object Dock: TDock97
    Left = 0
    Top = 263
    Width = 532
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 532
      Caption = 'Actions'
      ClientAreaHeight = 36
      ClientAreaWidth = 532
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAnnuler: TToolbarButton97
        Left = 120
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Annuler la derni'#232're action'
        Alignment = taRightJustify
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0075_S16G1'
      end
      object BFirst: TToolbarButton97
        Left = 4
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Premier'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFirstClick
        GlobalIndexImage = 'Z0301_S16G1'
      end
      object BPrev: TToolbarButton97
        Left = 32
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Pr'#233'c'#233'dent'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPrevClick
        GlobalIndexImage = 'Z0017_S16G1'
      end
      object BNext: TToolbarButton97
        Left = 60
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Suivant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNextClick
        GlobalIndexImage = 'Z0031_S16G1'
      end
      object BLast: TToolbarButton97
        Left = 88
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Dernier'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BLastClick
        GlobalIndexImage = 'Z0264_S16G1'
      end
      object BInsert: TToolbarButton97
        Left = 152
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsertClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 184
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 485
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BFerme: TToolbarButton97
        Left = 453
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 422
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Valider'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BImprimer: TToolbarButton97
        Left = 390
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BSousPlanBud: TToolbarButton97
        Left = 216
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Param'#233'trage des sous-plans budg'#233'taires'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSousPlanBudClick
        GlobalIndexImage = 'Z0008_S16G1'
      end
    end
  end
  object DBNav: TDBNavigator
    Left = 220
    Top = 108
    Width = 80
    Height = 18
    DataSource = Sa
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object FAutoSave: TCheckBox
    Left = 212
    Top = 68
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
  object HM: THMsgBox
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
      '2;?caption?;Vous devez renseigner un code !;W;O;O;O;'
      '3;?caption?;Vous devez renseigner un libell'#233' !;W;O;O;O;'
      
        '4;?caption?;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez ' +
        'le modifier;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '6;?caption?;Le code est incorrect. V'#233'rifiez qu'#39'il n'#39'est ni vide,' +
        ' ni renseign'#233' avec un code d'#233'j'#224' existant.;W;O;O;O;'
      
        '7;?caption?;ATTENTION.Les budgets rattach'#233's '#224' ce code vont '#234'tre ' +
        'mis '#224' jour. D'#233'sirez vous continuer?;Q;YNC;N;N;'
      
        '8;?caption?;Nombre maximum d'#39'enregistrement atteint. Vous ne pou' +
        'vez plus cr'#233'er de nouvel enregistrement.;W;O;O;O;'
      
        '9;?caption?;Vous devez param'#233'trer les sous plans des cat'#233'gories!' +
        ';W;O;O;O;')
    Left = 16
    Top = 124
  end
  object Ta: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TaAfterPost
    AfterDelete = TaAfterDelete
    OnNewRecord = TaNewRecord
    OnPostError = TaPostError
    EnableBCD = False
    IndexName = 'CC_CLE1'
    TableName = 'CHOIXCOD'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 48
    Top = 124
    object TaCC_TYPE: TStringField
      FieldName = 'CC_TYPE'
      Size = 3
    end
    object TaCC_CODE: TStringField
      FieldName = 'CC_CODE'
      EditMask = '>aaa;0; '
      Size = 3
    end
    object TaCC_LIBELLE: TStringField
      FieldName = 'CC_LIBELLE'
      Size = 35
    end
    object TaCC_ABREGE: TStringField
      FieldName = 'CC_ABREGE'
      Size = 17
    end
    object TaCC_LIBRE: TStringField
      FieldName = 'CC_LIBRE'
      Size = 70
    end
  end
  object Sa: TDataSource
    DataSet = Ta
    OnDataChange = SaDataChange
    OnUpdateData = SaUpdateData
    Left = 84
    Top = 124
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 120
    Top = 124
  end
end

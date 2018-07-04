object FStructure: TFStructure
  Left = 351
  Top = 149
  HelpContext = 1445000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Structures analytiques'
  ClientHeight = 347
  ClientWidth = 592
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
  object Panel3: TPanel
    Left = 273
    Top = 43
    Width = 319
    Height = 265
    Align = alClient
    BevelOuter = bvNone
    Caption = 'Panel3'
    TabOrder = 1
    object FListe2: THDBGrid
      Left = 0
      Top = 27
      Width = 319
      Height = 238
      Align = alClient
      DataSource = SSSPlan
      Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnKeyDown = FListe2KeyDown
      OnRowEnter = FListe2RowEnter
      Row = 1
      MultiSelection = False
      SortEnabled = False
      MyDefaultRowHeight = 0
      Columns = <
        item
          Expanded = False
          FieldName = 'PS_CODE'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ReadOnly = True
          Title.Caption = 'Code'
          Width = 76
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'PS_LIBELLE'
          Title.Caption = 'Libell'#233
          Width = 208
          Visible = True
        end>
    end
    object PPlan: TPanel
      Left = 0
      Top = 0
      Width = 319
      Height = 27
      Align = alTop
      Caption = 'Contenu du plan : '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 43
    Width = 273
    Height = 265
    Align = alLeft
    BevelOuter = bvNone
    Caption = 'Panel2'
    TabOrder = 5
    object FListe1: THDBGrid
      Left = 0
      Top = 27
      Width = 273
      Height = 238
      Align = alClient
      DataSource = SPlan
      Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnDblClick = FListe1DblClick
      OnKeyDown = FListe1KeyDown
      OnKeyPress = FListe1KeyPress
      OnRowEnter = FListe1RowEnter
      Row = 1
      MultiSelection = False
      SortEnabled = False
      MyDefaultRowHeight = 0
      Columns = <
        item
          Expanded = False
          FieldName = 'SS_SOUSSECTION'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ReadOnly = True
          Title.Caption = 'Plan'
          Width = 27
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'SS_LIBELLE'
          Title.Caption = 'Libell'#233
          Width = 127
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'SS_CONTROLE'
          Title.Caption = 'U.'
          Width = 14
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'SS_DEBUT'
          Title.Caption = 'D'#233'but'
          Width = 34
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'SS_LONGUEUR'
          Title.Caption = 'Long'
          Width = 31
          Visible = True
        end>
    end
    object Panel5: TPanel
      Left = 0
      Top = 0
      Width = 273
      Height = 27
      Align = alTop
      Caption = 'Plans de sous-sections:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
  end
  object PAxe: TTabControl
    Left = 0
    Top = 0
    Width = 592
    Height = 43
    Align = alTop
    TabOrder = 0
    Tabs.Strings = (
      'Axe n'#176'1'
      'Axe n'#176'2'
      'Axe n'#176'3'
      'Axe n'#176'4'
      'Axe n'#176'5')
    TabIndex = 0
    OnChange = PAxeChange
    object Lgaxe: TLabel
      Left = 5
      Top = 23
      Width = 72
      Height = 14
      AutoSize = False
      Caption = 'nn Caracteres'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object FAutoSave: TCheckBox
      Left = 276
      Top = 4
      Width = 45
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
      TabOrder = 0
      Visible = False
    end
  end
  object DBNav1: TDBNavigator
    Left = 145
    Top = 244
    Width = 80
    Height = 18
    DataSource = SPlan
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object DBNav2: TDBNavigator
    Left = 345
    Top = 180
    Width = 80
    Height = 18
    DataSource = SSSPlan
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 3
    Visible = False
  end
  object Dock: TDock97
    Left = 0
    Top = 308
    Width = 592
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 592
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 592
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BInsert2: TToolbarButton97
        Left = 303
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouvelle cat'#233'gorie'
        Caption = 'Nouvelle cat'#233'gorie'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsert2Click
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete2: TToolbarButton97
        Left = 335
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDelete2Click
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BDefaire2: TToolbarButton97
        Left = 271
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Annuler la derni'#232're action'
        Caption = 'Annuler contenu'
        DisplayMode = dmGlyphOnly
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
        OnClick = BDefaire2Click
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object BValider2: TToolbarButton97
        Left = 495
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider le contenu'
        Caption = 'Valider'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValider2Click
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 527
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object BAide2: TToolbarButton97
        Left = 559
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAide2Click
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BAutomate: TToolbarButton97
        Left = 367
        Top = 3
        Width = 28
        Height = 27
        Hint = 'G'#233'n'#233'ration automatique du contenu des plans'
        Caption = 'G'#233'n'#233'rer'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAutomateClick
        GlobalIndexImage = 'Z0775_S16G1'
      end
      object BMultiSect: TToolbarButton97
        Left = 399
        Top = 3
        Width = 28
        Height = 27
        Hint = 
          'G'#233'n'#233'ration des sections analytiques incluant le sous plan s'#233'lect' +
          'ionn'#233
        Caption = 'G'#233'n'#233'ration sections'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BMultiSectClick
        GlobalIndexImage = 'Z0192_S16G1'
      end
      object BSimpSect: TToolbarButton97
        Left = 431
        Top = 3
        Width = 28
        Height = 27
        Hint = 'G'#233'n'#233'ration simple des sections analytiques'
        Caption = 'G'#233'n'#233'ration simple'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSimpSectClick
        GlobalIndexImage = 'Z0210_S16G1'
      end
      object BDefaire1: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Annuler la derni'#232're action sur le plan'
        Caption = 'Annuler plan'
        DisplayMode = dmGlyphOnly
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
        OnClick = BDefaire1Click
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object BInsert1: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouveau plan'
        Caption = 'Nouveau plan'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsert1Click
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete1: TToolbarButton97
        Left = 68
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDelete1Click
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BValider1: TToolbarButton97
        Left = 100
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider le plan'
        Caption = 'Valider le plan'
        DisplayMode = dmGlyphOnly
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BValider1Click
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BImprimer1: TToolbarButton97
        Left = 133
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer le plan des sous-sections'
        Caption = 'Imprimer le plan des sous-sections'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        OnClick = BImprimer1Click
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BImprimer2: TToolbarButton97
        Left = 463
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer le contenu d'#39'un sous plan'
        Caption = 'Imprimer le contenu d'#39'un sous plan'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        OnClick = BImprimer2Click
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
  end
  object TPlan: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TPlanAfterPost
    BeforeDelete = TPlanBeforeDelete
    AfterDelete = TPlanAfterDelete
    OnNewRecord = TPlanNewRecord
    OnPostError = TPlanPostError
    EnableBCD = False
    IndexName = 'SS_CLE1'
    TableName = 'STRUCRSE'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 28
    Top = 232
    object TPlanSS_AXE: TStringField
      FieldName = 'SS_AXE'
      Visible = False
      Size = 3
    end
    object TPlanSS_SOUSSECTION: TStringField
      FieldName = 'SS_SOUSSECTION'
      EditMask = '>aaa;0; '
      Size = 3
    end
    object TPlanSS_LIBELLE: TStringField
      FieldName = 'SS_LIBELLE'
      Size = 35
    end
    object TPlanSS_CONTROLE: TStringField
      FieldName = 'SS_CONTROLE'
      Size = 1
    end
    object TPlanSS_DEBUT: TIntegerField
      FieldName = 'SS_DEBUT'
    end
    object TPlanSS_LONGUEUR: TIntegerField
      FieldName = 'SS_LONGUEUR'
    end
    object TPlanSS_SOCIETE: TStringField
      FieldName = 'SS_SOCIETE'
      Size = 3
    end
  end
  object SPlan: TDataSource
    DataSet = TPlan
    OnStateChange = SPlanStateChange
    OnDataChange = SPlanDataChange
    OnUpdateData = SPlanUpdateData
    Left = 80
    Top = 232
  end
  object TSSPlan: THTable
    MarshalOptions = moMarshalModifiedOnly
    OnNewRecord = TSSPlanNewRecord
    OnPostError = TSSPlanPostError
    EnableBCD = False
    IndexName = 'PS_CLE1'
    MasterFields = 'SS_AXE;SS_SOUSSECTION'
    MasterSource = SPlan
    TableName = 'SSSTRUCR'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 304
    Top = 228
    object TSSPlanPS_AXE: TStringField
      FieldName = 'PS_AXE'
      Visible = False
      Size = 3
    end
    object TSSPlanPS_SOUSSECTION: TStringField
      FieldName = 'PS_SOUSSECTION'
      Visible = False
      Size = 3
    end
    object TSSPlanPS_CODE: TStringField
      FieldName = 'PS_CODE'
      Size = 17
    end
    object TSSPlanPS_LIBELLE: TStringField
      FieldName = 'PS_LIBELLE'
      Size = 35
    end
    object TSSPlanPS_ABREGE: TStringField
      FieldName = 'PS_ABREGE'
      Size = 17
    end
    object TSSPlanPS_SOCIETE: TStringField
      FieldName = 'PS_SOCIETE'
      Size = 3
    end
  end
  object SSSPlan: TDataSource
    DataSet = TSSPlan
    OnStateChange = SSSPlanStateChange
    OnDataChange = SSSPlanDataChange
    OnUpdateData = SSSPlanUpdateData
    Left = 356
    Top = 228
  end
  object MsgBox1: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Structures analytiques;Voulez-vous enregistrer les modificatio' +
        'ns ?;Q;YNC;Y;C;'
      
        '1;Structures analytiques;Confirmez-vous la suppression de l'#39'enre' +
        'gistrement ?;Q;YNC;N;C;'
      '2;Structures analytiques;Vous devez renseigner un code.;W;O;O;O;'
      
        '3;Structures analytiques;Vous devez renseigner un libell'#233'.;W;O;O' +
        ';O;'
      
        '4;Structures analytiques;Le code que vous avez saisi existe d'#233'j'#224 +
        '. Vous devez le modifier.;W;O;O;O;'
      
        '5;Structures analytiques;Vous avez d'#233'j'#224' atteint la longueur maxi' +
        'mum des sections de cet axe.;W;O;O;O;'
      
        '6;Structures analytiques;La longueur que vous avez renseign'#233'e es' +
        't trop grande.;W;O;O;O;'
      
        '7;Structures analytiques;La modification de la longueur va entra' +
        #238'ner la destruction du contenu des plans de sous-sections. D'#233'sir' +
        'ez-vous continuez ?;Q;YNC;N;C;'
      
        '8;Structures analytiques;Votre axe est mal structur'#233' : il n'#39'y a ' +
        'pas de continuit'#233' dans les encha'#238'nements d'#233'but longueur.;W;O;O;O' +
        ';'
      
        '9;Structures analytiques;Une des longueurs que vous avez renseig' +
        'n'#233'es n'#39'est pas valide.;W;O;O;O;'
      
        '10;Structures analytiques;Vous devez renseigner un code.;W;O;O;O' +
        ';'
      
        '11;Structures analytiques;Vous ne pouvez pas renseigner un longu' +
        'eur '#233'gale '#224' 0.;W;O;O;O;'
      
        '12;Structures analytiques;Vous ne pouvez pas renseigner plus de ' +
        '3 sous-plans.;W;O;O;O;')
    Left = 88
    Top = 120
  end
  object MsgBox2: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Structures analytiques;Voulez-vous enregistrer les modificatio' +
        'ns ?;Q;YNC;Y;C;'
      
        '1;Structures analytiques;Confirmez-vous la suppression de l'#39'enre' +
        'gistrement ?;Q;YNC;N;C;'
      '2;Structures analytiques;Vous devez renseigner un code.;W;O;O;O;'
      
        '3;Structures analytiques;Vous devez renseigner un libell'#233'.;W;O;O' +
        ';O;'
      
        '4;Structures analytiques;Le code que vous avez saisi existe d'#233'j'#224 +
        '. Vous devez le modifier.;W;O;O;O;'
      
        '5;Structures analytiques;D'#233'sirez-vous g'#233'n'#233'rer les codes de sous-' +
        'sections '#224' partir des sections analytiques d'#233'j'#224' d'#233'finies ?;Q;YNC' +
        ';Y;C;'
      
        '6;Structures analytiques;Le code choisi ne respecte pas la longu' +
        'eur de cette sous-section.;W;O;O;O;'
      'G'#233'n'#233'ration des codes'
      
        '8;Structures analytiques;D'#233'sirez-vous g'#233'n'#233'rer toutes les combina' +
        'isons des sections analytiques incluant le sous plan s'#233'lectionn'#233 +
        ' ?;Q;ON;N;N;'
      '')
    Left = 380
    Top = 116
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Structures analytiques;Cet axe analytique n'#39'est pas structur'#233'.' +
        ';W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      'Contenu du plan : '
      'Caract'#232'res'
      'Plans de sous-sections:'
      '5;')
    Left = 212
    Top = 121
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 128
    Top = 184
  end
end

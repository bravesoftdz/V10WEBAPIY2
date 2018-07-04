object FVisuExp: TFVisuExp
  Left = 85
  Top = 150
  Width = 635
  Height = 300
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Liste des mouvements export'#233's'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PanelBouton: TPanel
    Left = 0
    Top = 238
    Width = 627
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    TabOrder = 0
    object BRechercher: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Rechercher dans la liste'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BRechercherClick
      GlobalIndexImage = 'Z0077_S16G1'
    end
    object Panel1: TPanel
      Left = 493
      Top = 2
      Width = 132
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 1
      object BImprimer: THBitBtn
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'BImprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BOuvrir: THBitBtn
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = FListeDblClick
        NumGlyphs = 2
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BAnnuler: THBitBtn
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: THBitBtn
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object FListe: THDBGrid
    Left = 0
    Top = 0
    Width = 627
    Height = 238
    Align = alClient
    DataSource = SEcr
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = FListeDblClick
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
    Columns = <
      item
        Expanded = False
        FieldName = 'E_JOURNAL'
        Title.Caption = 'Jal'
        Width = 25
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_DATECOMPTABLE'
        Title.Caption = 'Date'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_NUMEROPIECE'
        Title.Caption = 'Pi'#232'ce'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_GENERAL'
        Title.Caption = 'G'#233'n'#233'ral'
        Width = 103
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_AUXILIAIRE'
        Title.Caption = 'Auxiliaire'
        Width = 92
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_DEBIT'
        Title.Caption = 'D'#233'bit'
        Width = 84
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_CREDIT'
        Title.Caption = 'Cr'#233'dit'
        Width = 89
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_NATUREPIECE'
        Title.Caption = 'Nature'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'E_QUALIFPIECE'
        Title.Caption = 'Type'
        Visible = True
      end>
  end
  object QEcr: THQuery
    AutoCalcFields = False
    Filtered = True
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      '')
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 144
    Top = 42
  end
  object SEcr: TDataSource
    DataSet = QEcr
    Left = 96
    Top = 88
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 188
    Top = 88
  end
  object HQuery1: THQuery
    AutoCalcFields = False
    Filtered = True
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      '')
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 144
    Top = 92
  end
end

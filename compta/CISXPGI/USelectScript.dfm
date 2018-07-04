object FSelScript: TFSelScript
  Left = 464
  Top = 187
  Width = 310
  Height = 345
  ActiveControl = Panel1
  Caption = 'S'#233'lectionnez un script'
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 302
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object ScriptSelect: TEdit
      Left = 224
      Top = 4
      Width = 121
      Height = 21
      TabOrder = 0
      Visible = False
    end
    object Panel3: TPanel
      Left = 0
      Top = 0
      Width = 302
      Height = 41
      Align = alTop
      TabOrder = 1
      object TDOMAINE: TLabel
        Left = 39
        Top = 10
        Width = 42
        Height = 13
        Caption = 'Domaine'
      end
      object Domaine: THValComboBox
        Left = 101
        Top = 6
        Width = 122
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        ComboWidth = 300
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 33
    Width = 302
    Height = 278
    Align = alClient
    BevelOuter = bvNone
    Caption = 'Panel2'
    TabOrder = 1
    object DBGrid1: TDBGrid
      Left = 0
      Top = 0
      Width = 302
      Height = 278
      Align = alClient
      BiDiMode = bdLeftToRight
      DataSource = SQ
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
      ParentBiDiMode = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnDblClick = DBGrid1DblClick
    end
  end
  object SQ: TDataSource
    DataSet = Q
    Left = 90
    Top = 236
  end
  object Q: TQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      'Select'
      '  GZIMPCORRESP."Tablename",'
      '  GZIMPCORRESP."Profile"'
      
        'From "C:\Documents and Settings\import\parametre\GZIMPCORRESP.db' +
        '"'
      'As GZIMPCORRESP')
    Left = 70
    Top = 83
    object QProfile: TStringField
      DisplayLabel = 'Profil'
      FieldName = 'Profile'
    end
    object QTablename: TStringField
      DisplayWidth = 20
      FieldName = 'Tablename'
      Size = 35
    end
  end
end

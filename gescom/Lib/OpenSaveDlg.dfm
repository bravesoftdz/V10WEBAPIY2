object OpenSaveDialog: TOpenSaveDialog
  Left = 410
  Top = 188
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'OpenSaveDialog'
  ClientHeight = 323
  ClientWidth = 387
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object TCode: THLabel
    Left = 0
    Top = 4
    Width = 25
    Height = 13
    Caption = '&Code'
    FocusControl = ed_code
  end
  object TLib: THLabel
    Left = 108
    Top = 4
    Width = 30
    Height = 13
    Caption = '&Libell'#233
    FocusControl = ed_lib
  end
  object HPanel1: THPanel
    Left = 0
    Top = 288
    Width = 387
    Height = 35
    Align = alBottom
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object BValid: TBitBtn
      Left = 291
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
      ModalResult = 1
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValidClick
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        1800000000000003000000000000000000000000000000000000008284008284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828484000084000000828400
        8284008284008284008284008284008284008284008284008284008284008284
        0082848400000082000082008400000082840082840082840082840082840082
        8400828400828400828400828400828484000000820000820000820000820084
        0000008284008284008284008284008284008284008284008284008284840000
        0082000082000082000082000082000082008400000082840082840082840082
        8400828400828400828484000000820000820000820000FF0000820000820000
        8200008200840000008284008284008284008284008284008284008200008200
        00820000FF0000828400FF000082000082000082008400000082840082840082
        8400828400828400828400FF0000820000FF0000828400828400828400FF0000
        820000820000820084000000828400828400828400828400828400828400FF00
        00828400828400828400828400828400FF000082000082000082008400000082
        8400828400828400828400828400828400828400828400828400828400828400
        828400FF00008200008200008200840000008284008284008284008284008284
        00828400828400828400828400828400828400828400FF000082000082000082
        0084000000828400828400828400828400828400828400828400828400828400
        828400828400828400FF00008200008200008200840000008284008284008284
        00828400828400828400828400828400828400828400828400828400FF000082
        0000820000820084000000828400828400828400828400828400828400828400
        828400828400828400828400828400FF00008200008200840000008284008284
        0082840082840082840082840082840082840082840082840082840082840082
        8400FF0000820000820000828400828400828400828400828400828400828400
        828400828400828400828400828400828400828400FF00008284}
    end
    object BDel: THBitBtn
      Left = 4
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Supprimer'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BDelClick
      GlobalIndexImage = 'Z0005_S16G1'
    end
    object BAnnuler: THBitBtn
      Left = 323
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Annuler'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 355
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      GlobalIndexImage = 'Z1117_S16G1'
    end
  end
  object G_Lst: THDBGrid
    Left = 0
    Top = 44
    Width = 387
    Height = 244
    Align = alBottom
    DataSource = SQ
    Options = [dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    ReadOnly = True
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = G_LstDblClick
    Row = 0
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
  end
  object ed_code: THCritMaskEdit
    Left = 0
    Top = 20
    Width = 105
    Height = 21
    CharCase = ecUpperCase
    TabOrder = 0
    TagDispatch = 0
  end
  object ed_lib: THCritMaskEdit
    Left = 108
    Top = 20
    Width = 265
    Height = 21
    TabOrder = 1
    TagDispatch = 0
  end
  object SQ: TDataSource
    OnDataChange = SQDataChange
    Left = 356
    Top = 256
  end
end

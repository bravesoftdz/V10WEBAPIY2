object FCPGetFileViewer: TFCPGetFileViewer
  Left = 311
  Top = 176
  Width = 696
  Height = 480
  Caption = 'Visualisation d'#39'un document'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 399
    Width = 688
    Height = 47
    AllowDrag = False
    Position = dpBottom
    object ToolWindow: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 43
      ClientWidth = 688
      Caption = 'ToolWindow'
      ClientAreaHeight = 43
      ClientAreaWidth = 688
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        688
        43)
      object btnHelp: TToolbarButton97
        Left = 657
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object btnAnnul: TToolbarButton97
        Left = 627
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Annuler'
        Anchors = [akTop, akRight]
        Flat = False
        ImageIndex = 3
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object btnAccept: TToolbarButton97
        Left = 596
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Accepter'
        Anchors = [akTop, akRight]
        Flat = False
        ModalResult = 1
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object btnOpen: TToolbarButton97
        Left = 449
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ouvrir une Image'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z1100_S16G1'
      end
      object bSauve: TToolbarButton97
        Tag = 21004
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Enregistrer sous'
        DisplayMode = dmGlyphOnly
        Caption = 'Sauvegarder'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bSauveClick
        GlobalIndexImage = 'Z0027_S16G1'
      end
      object btnCopy: TToolbarButton97
        Tag = 21004
        Left = 34
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Copier dans le presse papier'
        DisplayMode = dmGlyphOnly
        Caption = 'Sauvegarder'
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = btnCopyClick
        GlobalIndexImage = 'Z0820_S16G2'
      end
      object btnPrint: TToolbarButton97
        Left = 64
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Impression document'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnPrintClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object btnPan: TToolbarButton97
        Left = 98
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Main'
        GroupIndex = 1
        Down = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnPanClick
        GlobalIndexImage = 'Z0995_S16G1'
      end
      object btnZoom: TToolbarButton97
        Left = 129
        Top = 2
        Width = 39
        Height = 27
        Hint = 'Zoom avant'
        GroupIndex = 1
        DropdownAlways = True
        DropdownArrow = True
        DropdownCombo = True
        DropdownMenu = PopupMenuZoom
        Flat = False
        ImageIndex = 0
        Images = ImageListZoom
        ParentShowHint = False
        ShowHint = True
        OnClick = btnZoomClick
      end
      object btnFitToPage: TToolbarButton97
        Left = 168
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Taille '#233'cran'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnFitToPageClick
        GlobalIndexImage = 'Z0039_S24G1'
      end
      object btnFitToLarge: TToolbarButton97
        Left = 198
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Pleine largeur'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnFitToLargeClick
        GlobalIndexImage = 'Z0042_S24G1'
      end
      object btnFitTo100: TToolbarButton97
        Left = 228
        Top = 2
        Width = 28
        Height = 27
        Hint = '100 %'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnFitTo100Click
        GlobalIndexImage = 'Z0039_S24G1'
      end
      object btnFirstPage: TToolbarButton97
        Left = 260
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Premi'#232're page'
        DisplayMode = dmGlyphOnly
        Caption = 'Premi'#232're page'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        OnClick = btnFirstPageClick
        GlobalIndexImage = 'Z0301_S16G1'
      end
      object btnPagePrevious: TToolbarButton97
        Left = 290
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Page pr'#233'c'#233'dente'
        DisplayMode = dmGlyphOnly
        Caption = 'Page pr'#233'c'#233'dente'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        OnClick = btnPagePreviousClick
        GlobalIndexImage = 'Z0017_S16G1'
      end
      object btnPageNext: TToolbarButton97
        Left = 380
        Top = 2
        Width = 30
        Height = 27
        Hint = 'Page suivante'
        DisplayMode = dmGlyphOnly
        Caption = 'Page suivante'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        OnClick = btnPageNextClick
        GlobalIndexImage = 'Z0031_S16G1'
      end
      object btnLastPage: TToolbarButton97
        Left = 410
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Derni'#232're page'
        DisplayMode = dmGlyphOnly
        Caption = 'Derni'#232're page'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        OnClick = btnLastPageClick
        GlobalIndexImage = 'Z0264_S16G1'
      end
      object BRotation90: TToolbarButton97
        Left = 481
        Top = 2
        Width = 28
        Height = 27
        Flat = False
        Glyph.Data = {
          D2000000424DD200000000000000520000002800000010000000100000000100
          040000000000800000001E1C00001E1C0000070000000700000000000000FFFF
          FF0000FCFF00007778008EFEFF00C3C8C800FFFFFF0055555533333555555355
          3322222335554233222222222355422222244442223542222445555422354222
          2355555542234222223555554443444444555555555555555555553333334333
          5555542222234223555555422223542235555332222354222333322222235542
          22222222442355544222224455455555544444555555}
        GlyphMask.Data = {00000000}
        OnClick = BRotation90Click
      end
      object ToolbarButton971: TToolbarButton97
        Left = 512
        Top = 2
        Width = 41
        Height = 27
        DropdownArrow = True
        DropdownMenu = PopUpRevInteg
        Glyph.Data = {
          06020000424D0602000000000000760000002800000028000000140000000100
          0400000000009001000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333FFFFFFFFFF333333333000000000033333333337777777777F3333
          333330FEFEFEFE033333333337F3FFFFFF7F3333333330E444444F0333333333
          37F77777737F333333333099999999033333333337F33333337F33333333309F
          FFFFF9033333333337F33333337F333333333099999999033333333337F3FFFF
          FF7F3333333330E444444F033333333337F77777737F3333333330FEFEFEFE03
          3333333337F3FFF3FF7F3333333330E444E000033333333337F7773777733333
          333330FEFEF0F0333333333337F33337F7333333333330EFEFE0033333333333
          37FFFFF7733333FFFF333000000033333000033337777777333337777F333333
          3333333330EA0333333333333333F7F37FFF33333333333000AE000333333333
          33377733777F333333333330EAEAEA03333333333337FFF33F7F333333333330
          00AE000333333333333777F3777333333333333330EA033333333333333337FF
          7F33333333333333300003333333333333333777733333333333333333333333
          33333333333333333333}
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
      end
      object panelInfoPages: THPanel
        Left = 320
        Top = 3
        Width = 58
        Height = 24
        Hint = 'Indicateur de pages'
        BevelOuter = bvNone
        BorderStyle = bsSingle
        Caption = '0/0'
        FullRepaint = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
    end
  end
  object PanelView: THPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 399
    Align = alClient
    FullRepaint = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object ImageListZoom: THImageList
    GlobalIndexImages.Strings = (
      'Z1027_S16G1'
      'Z1902_S16G1')
    Left = 152
    Top = 144
  end
  object PopupMenuZoom: THPopupMenu
    Images = ImageListZoom
    Left = 152
    Top = 192
    object ZoomForward: THMenuItem
      Caption = 'Zoom avant'
      ImageIndex = 0
      OnClick = ZoomForwardClick
    end
    object ZoomBackward: THMenuItem
      Caption = 'Zoom arri'#232're'
      ImageIndex = 1
      OnClick = ZoomBackwardClick
    end
  end
  object OpenDialog: THOpenDialog
    Left = 152
    Top = 232
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 103
    Top = 78
  end
  object PopUpRevInteg: THPopupMenu
    OnPopup = PopUpRevIntegPopup
    Left = 440
    Top = 128
    object DOCUMENTATIONTRAVAUX: TMenuItem
      Caption = 'Documentation des travau&x'
      OnClick = DOCUMENTATIONTRAVAUXClick
      object MEMOCYCLE: TMenuItem
        Caption = 'Commentaire sur le c&ycle'
        OnClick = MEMOCYCLEClick
      end
      object MEMOOBJECTIF: TMenuItem
        Caption = '&Objectif de r'#233'vision'
        OnClick = MEMOOBJECTIFClick
      end
      object MEMOSYNTHESE: TMenuItem
        Caption = '&Synth'#232'se du cycle'
        OnClick = MEMOSYNTHESEClick
      end
      object MEMOMILLESIME: TMenuItem
        Caption = 'Commentaire &mill'#233'sim'#233' du compte'
        ShortCut = 32847
        OnClick = MEMOMILLESIMEClick
      end
      object MEMOCOMPTE: TMenuItem
        Caption = 'Commentaire sur le &compte'
        OnClick = MEMOCOMPTEClick
      end
    end
    object INFOCOMP: TMenuItem
      Caption = '&Informations compl'#233'mentaires'
      OnClick = INFOCOMPClick
    end
  end
end

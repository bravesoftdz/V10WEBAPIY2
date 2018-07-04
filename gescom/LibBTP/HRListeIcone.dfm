object FicheIcone: TFicheIcone
  Left = 304
  Top = 196
  Width = 373
  Height = 308
  Caption = 'Choisir un ic'#244'ne'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LIcone: TListView
    Left = 0
    Top = 0
    Width = 365
    Height = 241
    Align = alClient
    Columns = <>
    IconOptions.AutoArrange = True
    LargeImages = LImage
    ShowWorkAreas = True
    TabOrder = 0
  end
  object Dock971: TDock97
    Left = 0
    Top = 241
    Width = 365
    Height = 33
    AllowDrag = False
    Position = dpBottom
    object Toolbar971: TToolbar97
      Left = 309
      Top = 0
      DockableTo = [dpTop, dpRight]
      DockPos = 309
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      DesignSize = (
        54
        29)
      object BAnnul: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 29
        Alignment = taRightJustify
        Anchors = [akRight, akBottom]
        Opaque = False
        OnClick = BAnnuClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object Bok: TToolbarButton97
        Left = 0
        Top = 0
        Width = 27
        Height = 29
        Alignment = taRightJustify
        Anchors = [akLeft, akTop, akRight, akBottom]
        NumGlyphs = 2
        Opaque = False
        OnClick = BOkClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
    end
  end
  object LImage: THImageList
    Height = 25
    Width = 25
    Left = 100
    Top = 104
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 228
    Top = 116
  end
end

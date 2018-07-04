object ViewerExt: TViewerExt
  Left = 199
  Top = 148
  HorzScrollBar.Visible = False
  AutoScroll = False
  Caption = 'Visualisation'
  ClientHeight = 412
  ClientWidth = 689
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  KeyPreview = True
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object SPBOTTOM: THSplitter
    Left = 0
    Top = 357
    Width = 689
    Height = 7
    Cursor = crVSplit
    Align = alBottom
    Visible = False
  end
  object SPRIGHT: THSplitter
    Left = 682
    Top = 10
    Width = 7
    Height = 341
    Align = alRight
    Visible = False
  end
  object SPTOP: THSplitter
    Left = 0
    Top = 0
    Width = 689
    Height = 8
    Cursor = crVSplit
    Align = alTop
    Visible = False
  end
  object SPLEFT: THSplitter
    Left = 0
    Top = 10
    Width = 8
    Height = 341
    Visible = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 364
    Width = 689
    Height = 48
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 44
      ClientWidth = 689
      ClientAreaHeight = 44
      ClientAreaWidth = 689
      DockPos = 0
      FullSize = True
      Resizable = False
      TabOrder = 0
      DesignSize = (
        689
        44)
      object BValider: TToolbarButton97
        Left = 585
        Top = 18
        Width = 28
        Height = 28
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
      end
      object BFerme: TToolbarButton97
        Left = 617
        Top = 18
        Width = 28
        Height = 28
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 649
        Top = 18
        Width = 28
        Height = 28
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BMENUVIEWER: TToolbarButton97
        Left = 529
        Top = 18
        Width = 36
        Height = 28
        Hint = 'Actions document'
        Anchors = [akRight, akBottom]
        DropdownArrow = True
        DropdownMenu = POPUPVIEWER
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0138_S16G1'
      end
    end
  end
  object PBOTTOM: THPanel
    Left = 0
    Top = 351
    Width = 689
    Height = 6
    Align = alBottom
    FullRepaint = False
    TabOrder = 1
    Visible = False
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object PTOP: THPanel
    Left = 0
    Top = 8
    Width = 689
    Height = 2
    Align = alTop
    FullRepaint = False
    TabOrder = 2
    Visible = False
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object PLEFT: THPanel
    Left = 8
    Top = 10
    Width = 1
    Height = 341
    Align = alLeft
    FullRepaint = False
    TabOrder = 3
    Visible = False
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object PRIGHT: THPanel
    Left = 681
    Top = 10
    Width = 1
    Height = 341
    Align = alRight
    FullRepaint = False
    TabOrder = 4
    Visible = False
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object POPUPVIEWER: TPopupMenu
    Left = 254
    Top = 108
    object VIEWERNEXT: TMenuItem
      Caption = 'Document suivant'
    end
    object VIEWERPREC: TMenuItem
      Caption = 'Document pr'#233'cedent'
    end
    object VIEWERMOVE: TMenuItem
      Caption = 'D'#233'placer vers'
      object VIEWERMOVEACH: TMenuItem
        Caption = 'Achat'
      end
      object VIEWERMOVEBQE: TMenuItem
        Caption = 'Banque'
      end
      object VIEWERMOVECAI: TMenuItem
        Caption = 'Caisse'
      end
      object VIEWERMOVEOD: TMenuItem
        Caption = 'OD'
      end
      object VIEWERMOVEVTE: TMenuItem
        Caption = 'Vente'
      end
      object VIEWERMOVECORBEILLE: TMenuItem
        Caption = 'Corbeille'
      end
    end
    object VIEWERPRINT: TMenuItem
      Caption = 'Imprimer le document'
    end
    object VIEWERPAN: TMenuItem
      Caption = 'Main'
    end
    object VIEWERZOOMAV: TMenuItem
      Caption = 'Zoom avant'
    end
    object VIEWERZOOMAR: TMenuItem
      Caption = 'Zoom arri'#232're'
    end
    object VIEWERFITPAGE: TMenuItem
      Caption = 'Taille '#233'cran'
    end
    object VIEWERFITLARGE: TMenuItem
      Caption = 'Pleine largeur'
    end
    object VIEWERFIT100: TMenuItem
      Caption = '100 %'
    end
    object VIEWERROTATE: TMenuItem
      Caption = 'Retourner le document'
    end
  end
end

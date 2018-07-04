object FListeSaisie: TFListeSaisie
  Left = 311
  Top = 227
  Width = 658
  Height = 424
  HelpContext = 110000031
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Listes de saisie'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 349
    Width = 642
    Height = 36
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 633
      Caption = 'Outils listes de saisie'
      CloseButton = False
      ClientAreaHeight = 32
      ClientAreaWidth = 633
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        633
        32)
      object BValider: TToolbarButton97
        Left = 536
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer'
        Caption = 'Enregistrer'
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
        Layout = blGlyphTop
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 568
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 600
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object tbLoad: TToolbarButton97
        Left = 38
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ouvrir liste'
        Caption = '&Ouvrir liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        OnClick = tbLoadClick
        GlobalIndexImage = 'Z0174_S16G1'
      end
      object tbNew: TToolbarButton97
        Left = 6
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouvelle liste'
        Caption = '&Nouvelle liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        OnClick = tbNewClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
    end
  end
  object G_Liste: THGrid
    Tag = 1
    Left = 0
    Top = 71
    Width = 642
    Height = 278
    Align = alClient
    ColCount = 9
    DefaultRowHeight = 27
    FixedCols = 0
    RowCount = 4
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goColMoving, goTabs, goThumbTracking]
    ScrollBars = ssVertical
    TabOrder = 1
    OnColumnMoved = G_ListeColumnMoved
    OnDragDrop = G_ListeDragDrop
    OnDragOver = G_ListeDragOver
    OnMouseUp = G_ListeMouseUp
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnCellEnter = G_ListeCellEnter
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ColWidths = (
      64
      64
      64
      64
      64
      64
      64
      64
      64)
  end
  object TDetailPrix: TToolWindow97
    Left = 508
    Top = 144
    ClientHeight = 167
    ClientWidth = 113
    Caption = 'Champs disponibles'
    CloseButton = False
    ClientAreaHeight = 167
    ClientAreaWidth = 113
    DockableTo = []
    HideWhenInactive = False
    Resizable = False
    TabOrder = 2
    object FieldsList: TListBox
      Left = 0
      Top = 0
      Width = 113
      Height = 167
      Align = alClient
      ItemHeight = 13
      TabOrder = 0
      OnEndDrag = FieldsListEndDrag
      OnMouseDown = FieldsListMouseDown
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 41
    Width = 642
    Height = 30
    object ToolWindow972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 26
      ClientWidth = 160
      Caption = 'Outils colonnes de liste'
      CloseButton = False
      ClientAreaHeight = 26
      ClientAreaWidth = 160
      DockPos = 0
      TabOrder = 0
      object tbDelCol: TToolbarButton97
        Left = 4
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Supprimer la colonne'
        Caption = '&Supprimer colonne'
        DisplayMode = dmGlyphOnly
        OnClick = tbDelColClick
        GlobalIndexImage = 'Z0375_S16G1'
      end
      object tbLeft: TToolbarButton97
        Left = 36
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Aligner '#224' gauche'
        Caption = 'Aligner '#224' &gauche'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        OnClick = tbLeftClick
        GlobalIndexImage = 'Z0280_S16G1'
      end
      object tbCenter: TToolbarButton97
        Left = 64
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Centrer'
        Caption = '&Centrer'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        OnClick = tbCenterClick
        GlobalIndexImage = 'Z0280_S16G1'
      end
      object tbRight: TToolbarButton97
        Left = 92
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Aligner '#224' droite'
        Caption = 'Aligner '#224' &droite'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        OnClick = tbRightClick
        GlobalIndexImage = 'Z0280_S16G1'
      end
      object SpinDec: TSpinEdit
        Left = 124
        Top = 2
        Width = 33
        Height = 22
        Hint = 'Nombre de d'#233'cimales'
        EditorEnabled = False
        MaxValue = 5
        MinValue = 0
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        Value = 0
        OnChange = SpinDecChange
      end
    end
  end
  object PEntete: THPanel
    Left = 0
    Top = 0
    Width = 642
    Height = 41
    Align = alTop
    FullRepaint = False
    TabOrder = 4
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 40
    Top = 244
  end
  object FindLigne: TFindDialog
    Left = 72
    Top = 244
  end
  object MsgCaptions: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Le nombre maximum de colonnes est '
      
        'Vous devez supprimer une colonne pour pouvoir en ins'#233'rer une aut' +
        're.'
      'Voulez-vous abandonner la liste en cours ?'
      'Cette liste existe d'#233'j'#224', voulez-vous la remplacer ?'
      'Etes-vous s'#251'r(e) de vouloir effacer la liste '
      'Ouvrir une liste de saisie'
      'Enregistrer la liste de saisie'
      'Le code liste doit commencer par GCSAISIE')
    Left = 108
    Top = 244
  end
end

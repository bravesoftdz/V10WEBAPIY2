object FParamSaisEff: TFParamSaisEff
  Left = 311
  Top = 227
  Width = 640
  Height = 400
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 
    'Param'#233'trage de la liste de saisie des effets en retour d'#39'accepta' +
    'tion'
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
    Top = 337
    Width = 632
    Height = 36
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 623
      Caption = 'Outils listes de saisie'
      CloseButton = False
      ClientAreaHeight = 32
      ClientAreaWidth = 623
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        623
        32)
      object BValider: TToolbarButton97
        Left = 526
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer'
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Enregistrer'
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
        Left = 558
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
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
        Left = 590
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
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
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object tbLoad: TToolbarButton97
        Left = 38
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ouvrir liste'
        DisplayMode = dmGlyphOnly
        Caption = '&Ouvrir liste'
        Flat = False
        Visible = False
        OnClick = tbLoadClick
        GlobalIndexImage = 'Z0174_S16G1'
      end
      object tbNew: TToolbarButton97
        Left = 6
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouvelle liste'
        DisplayMode = dmGlyphOnly
        Caption = '&Nouvelle liste'
        Flat = False
        Visible = False
        OnClick = tbNewClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
    end
  end
  object G_Liste: THGrid
    Tag = 1
    Left = 0
    Top = 30
    Width = 632
    Height = 307
    Align = alClient
    ColCount = 9
    DefaultRowHeight = 27
    FixedCols = 0
    RowCount = 3
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
    TwoColors = False
    AlternateColor = 13224395
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
    Left = 248
    Top = 136
    ClientHeight = 167
    ClientWidth = 116
    Caption = 'Champs disponibles'
    CloseButton = False
    ClientAreaHeight = 167
    ClientAreaWidth = 116
    DockableTo = []
    HideWhenInactive = False
    TabOrder = 2
    object FieldsList: TListBox
      Left = 0
      Top = 0
      Width = 116
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
    Top = 0
    Width = 632
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
        DisplayMode = dmGlyphOnly
        Caption = '&Supprimer colonne'
        OnClick = tbDelColClick
        GlobalIndexImage = 'Z0375_S16G1'
      end
      object tbLeft: TToolbarButton97
        Left = 36
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Aligner '#224' gauche'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = 'Aligner '#224' &gauche'
        OnClick = tbLeftClick
        GlobalIndexImage = 'Z0280_S16G1'
      end
      object tbCenter: TToolbarButton97
        Left = 64
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Centrer'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = '&Centrer'
        OnClick = tbCenterClick
        GlobalIndexImage = 'Z0280_S16G1'
      end
      object tbRight: TToolbarButton97
        Left = 92
        Top = 2
        Width = 24
        Height = 21
        Hint = 'Aligner '#224' droite'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = 'Aligner '#224' &droite'
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
  object HM: THMsgBox
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
      'Confirmez-vous la validation du param'#233'trage ?'
      'Ouvrir une liste de saisie'
      'Enregistrer la liste de saisie')
    Left = 108
    Top = 244
  end
end

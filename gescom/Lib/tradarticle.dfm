object FTradArticle: TFTradArticle
  Left = 28
  Top = 54
  Width = 623
  Height = 446
  HelpContext = 110000082
  Caption = 'Traduction libell'#233's articles'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PLANGUE: THPanel
    Left = 0
    Top = 0
    Width = 615
    Height = 43
    Align = alTop
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGTA_LANGUE: THLabel
      Left = 20
      Top = 15
      Width = 36
      Height = 13
      Caption = '&Langue'
      FocusControl = GTA_LANGUE
    end
    object GTA_LANGUE: THValComboBox
      Left = 79
      Top = 11
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = GTA_LANGUEChange
      TagDispatch = 0
      DataType = 'TTLANGUE'
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 376
    Width = 615
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object Toolbar972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 615
      Caption = 'Outils Tarif'
      ClientAreaHeight = 32
      ClientAreaWidth = 615
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 509
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
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
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 541
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
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
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 572
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
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
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 57
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher '
        DisplayMode = dmGlyphOnly
        Caption = 'Rechercher'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 10
        Top = 3
        Width = 40
        Height = 27
        Hint = 'Voir article'
        DisplayMode = dmGlyphOnly
        Caption = 'Compl'#233'ments'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BInfosClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BInsert: TToolbarButton97
        Tag = 1
        Left = 92
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Caption = 'Compl'#233'ments'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BInsertClick
        GlobalIndexImage = 'Z0053_S16G1'
        IsControl = True
      end
      object BDelete: TToolbarButton97
        Tag = 1
        Left = 126
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Compl'#233'ments'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
        IsControl = True
      end
    end
  end
  object PPIED: THPanel
    Left = 0
    Top = 333
    Width = 615
    Height = 43
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGA_GRILLEDIM1: THLabel
      Left = 16
      Top = 4
      Width = 92
      Height = 13
      Caption = 'TGA_GRILLEDIM1'
    end
    object TGA_GRILLEDIM2: THLabel
      Left = 136
      Top = 4
      Width = 92
      Height = 13
      Caption = 'TGA_GRILLEDIM2'
    end
    object TGA_GRILLEDIM3: THLabel
      Left = 256
      Top = 4
      Width = 92
      Height = 13
      Caption = 'TGA_GRILLEDIM3'
    end
    object TGA_GRILLEDIM4: THLabel
      Left = 376
      Top = 4
      Width = 92
      Height = 13
      Caption = 'TGA_GRILLEDIM4'
    end
    object TGA_GRILLEDIM5: THLabel
      Left = 496
      Top = 4
      Width = 92
      Height = 13
      Caption = 'TGA_GRILLEDIM5'
    end
    object TGA_CODEDIM1: THCritMaskEdit
      Left = 13
      Top = 19
      Width = 109
      Height = 21
      TabOrder = 0
      Text = 'TGA_CODEDIM1'
      TagDispatch = 0
      Libelle = TGA_GRILLEDIM1
    end
    object TGA_CODEDIM2: THCritMaskEdit
      Left = 133
      Top = 19
      Width = 109
      Height = 21
      TabOrder = 1
      Text = 'TGA_CODEDIM2'
      TagDispatch = 0
      Libelle = TGA_GRILLEDIM2
    end
    object TGA_CODEDIM3: THCritMaskEdit
      Left = 253
      Top = 19
      Width = 109
      Height = 21
      TabOrder = 2
      Text = 'TGA_CODEDIM3'
      TagDispatch = 0
      Libelle = TGA_GRILLEDIM3
    end
    object TGA_CODEDIM4: THCritMaskEdit
      Left = 373
      Top = 19
      Width = 109
      Height = 21
      TabOrder = 3
      Text = 'TGA_CODEDIM4'
      TagDispatch = 0
      Libelle = TGA_GRILLEDIM4
    end
    object TGA_CODEDIM5: THCritMaskEdit
      Left = 493
      Top = 19
      Width = 109
      Height = 21
      TabOrder = 4
      Text = 'TGA_CODEDIM5'
      TagDispatch = 0
      Libelle = TGA_GRILLEDIM5
    end
  end
  object PGRID: THPanel
    Left = 0
    Top = 43
    Width = 615
    Height = 290
    Align = alClient
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object GTrad: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 613
      Height = 288
      Align = alClient
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goTabs]
      TabOrder = 0
      OnEnter = GTradEnter
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GTradRowEnter
      OnRowExit = GTradRowExit
      OnCellEnter = GTradCellEnter
      OnCellExit = GTradCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = GTradElipsisClick
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 488
    Top = 16
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      '1;?caption?;Voulez-vous supprimer l'#39'enregistrement ?;Q;YNC;N;C;'
      '2;?caption?;Cet article n'#39'existe pas.;W;O;O;O;'
      '3;?caption?;La traduction est d'#233'j'#224' faite.;W;O;O;O;'
      '4;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;Y;N;')
    Left = 466
    Top = 204
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 440
    Top = 16
  end
  object HMess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ATTENTION : Traduction non enregistr'#233'e !'
      
        'ATTENTION : Cette traduction, en cours de traitement par un autr' +
        'e utilisateur, n'#39'a pas '#233't'#233' enregistr'#233'e !')
    Left = 424
    Top = 204
  end
end

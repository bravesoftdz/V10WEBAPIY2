object FResulSuggestion: TFResulSuggestion
  Left = 95
  Top = 170
  Width = 696
  Height = 480
  Align = alClient
  BorderIcons = [biSystemMenu]
  Caption = 'Resultat d'#39'une suggestion'
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object G_Fin: THGrid
    Left = 172
    Top = 312
    Width = 490
    Height = 96
    ColCount = 15
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 15
    TabOrder = 5
    Visible = False
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
  end
  object G_Evol: THGrid
    Tag = 1
    Left = 169
    Top = 210
    Width = 490
    Height = 96
    ColCount = 15
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 15
    TabOrder = 4
    Visible = False
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
  end
  object G_Mvts: THGrid
    Left = 172
    Top = 108
    Width = 490
    Height = 96
    ColCount = 15
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 15
    TabOrder = 3
    Visible = False
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
  end
  object G_LIG: THGrid
    Left = 0
    Top = 105
    Width = 680
    Height = 259
    Align = alClient
    ColCount = 8
    DefaultRowHeight = 18
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goTabs, goAlwaysShowEditor]
    TabOrder = 0
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = G_LIGRowEnter
    OnCellEnter = G_LIGCellEnter
    OnCellExit = G_LIGCellExit
    ColCombo = 5
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ElipsisButton = True
    OnElipsisClick = G_LIGElipsisClick
  end
  object PANTETE: TPanel
    Left = 0
    Top = 0
    Width = 680
    Height = 105
    Align = alTop
    TabOrder = 1
    object TSR_CREATION: THLabel
      Left = 52
      Top = 11
      Width = 69
      Height = 13
      Caption = 'Date de calcul'
      FocusControl = GZZ_CREATION
    end
    object BZoomArticle: TToolbarButton97
      Tag = 100
      Left = 8
      Top = 36
      Width = 28
      Height = 27
      Hint = 'Voir article'
      Caption = 'Article'
      Enabled = False
      Visible = False
      OnClick = BZoomArticleClick
    end
    object BZoomTiers: TToolbarButton97
      Tag = 100
      Left = 44
      Top = 36
      Width = 28
      Height = 27
      Hint = 'Voir fournisseur'
      Caption = 'Fournisseur'
      Enabled = False
      Visible = False
      OnClick = BZoomTiersClick
    end
    object BZoomLigne: TToolbarButton97
      Tag = 100
      Left = 76
      Top = 36
      Width = 28
      Height = 27
      Hint = 'Infos Ligne'
      Caption = 'Info Ligne'
      Enabled = False
      Visible = False
      OnClick = BZoomLigneClick
    end
    object Recap: TColorMemo
      Left = 376
      Top = 7
      Width = 303
      Height = 89
      Color = clBtnFace
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 0
      ColorRules = <>
    end
    object GZZ_CREATION: THCritMaskEdit
      Left = 157
      Top = 7
      Width = 86
      Height = 21
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
      TagDispatch = 0
      OpeType = otDate
      ControlerDate = True
    end
    object GZZ_SUGGEST: THCritMaskEdit
      Left = 160
      Top = 48
      Width = 121
      Height = 21
      Color = clYellow
      TabOrder = 2
      Visible = False
      OnExit = GZZ_SUGGESTExit
      TagDispatch = 0
    end
  end
  object DetailFourni: TToolWindow97
    Left = 204
    Top = 140
    ClientHeight = 138
    ClientWidth = 409
    BorderStyle = bsNone
    Caption = 'D'#233'tail des crit'#232'res de selection du fournisseur'
    ClientAreaHeight = 138
    ClientAreaWidth = 409
    TabOrder = 6
    Visible = False
    object G_FOU: THGrid
      Left = 0
      Top = 0
      Width = 409
      Height = 138
      Align = alClient
      DefaultRowHeight = 16
      FixedCols = 0
      RowCount = 10
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
      TabOrder = 0
      OnDblClick = G_FOUDblClick
      SortedCol = -1
      Titres.Strings = (
        'Fournisseur'
        'R'#233'f'#233'rence'
        'Dernier PA'
        'Cote'
        'Delai livraison')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      OnSorted = G_FOUSorted
      ColWidths = (
        109
        78
        75
        49
        72)
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 405
    Width = 680
    Height = 36
    AllowDrag = False
    LimitToOneRow = True
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 680
      Caption = 'ToolWindow971'
      ClientAreaHeight = 32
      ClientAreaWidth = 680
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object bDetail: TToolbarButton97
        Left = 45
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'tail des donn'#233'es de l'#39'article'
        Caption = 'D'#233'tail'
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
        OnClick = bDetailClick
        GlobalIndexImage = 'Z0202_S16G1'
      end
      object bTri: TToolbarButton97
        Left = 75
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'tail des crit'#232'res de choix du fournisseur'
        Caption = 'D'#233'tail'
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
        OnClick = bTriClick
        GlobalIndexImage = 'Z0202_S16G1'
      end
      object bRecalcul: TToolbarButton97
        Left = 105
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Recalcul des priorit'#233's'
        Caption = 'D'#233'tail'
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
        OnClick = bRecalculClick
        GlobalIndexImage = 'Z0050_S16G1'
      end
      object bGenPiece: TToolbarButton97
        Left = 142
        Top = 2
        Width = 28
        Height = 27
        Hint = 'G'#233'n'#233'ration d'#39'une pi'#233'ce d'#39'achat'
        Caption = 'D'#233'tail'
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
        Visible = False
        OnClick = bRecalculClick
        GlobalIndexImage = 'Z0222_S16G1'
      end
      object bPrint: TToolbarButton97
        Left = 559
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bPrintClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 649
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 11000305
        Caption = 'Aide'
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
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 619
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
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
      object BValider: TToolbarButton97
        Left = 589
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Enregistrer '
        Caption = 'Enregistrer'
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
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 9
        Top = 2
        Width = 36
        Height = 27
        Hint = 'Menu zoom'
        Caption = 'Zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 172
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer la pi'#232'ce'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 201
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la pi'#232'ce'
        Caption = 'Rechercher'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
    end
  end
  object DetailEvol: TToolWindow97
    Left = 52
    Top = 176
    ClientHeight = 193
    ClientWidth = 453
    BorderStyle = bsNone
    Caption = 'Evolution du stock'
    ClientAreaHeight = 193
    ClientAreaWidth = 453
    DockableTo = [dpTop, dpBottom]
    MinClientWidth = 557
    TabOrder = 2
    Visible = False
    object G_DET: THGrid
      Left = 0
      Top = 25
      Width = 453
      Height = 168
      Align = alClient
      ColCount = 6
      DefaultRowHeight = 16
      FixedCols = 0
      RowCount = 25
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Date'
        'Pi'#232'ce'
        'Stk physique'
        'R'#233'serv'#233' Client'
        'Attendu fourn.'
        'Disponible')
      Couleur = False
      MultiSelect = True
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ColWidths = (
        67
        64
        74
        77
        73
        74)
    end
    object DataLigne: TMemo
      Left = 4
      Top = 100
      Width = 185
      Height = 89
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Lines.Strings = (
        'DataLigne')
      ParentFont = False
      TabOrder = 1
      Visible = False
      WantReturns = False
      WordWrap = False
      OnChange = DataLigneChange
    end
    object Panel1: TPanel
      Left = 0
      Top = 0
      Width = 453
      Height = 25
      Align = alTop
      TabOrder = 2
      object LStockMin: TLabel
        Left = 8
        Top = 6
        Width = 3
        Height = 13
      end
      object LStockMax: TLabel
        Left = 224
        Top = 6
        Width = 3
        Height = 13
      end
    end
  end
  object Pbas: TPanel
    Left = 0
    Top = 364
    Width = 680
    Height = 41
    Align = alBottom
    TabOrder = 8
    object PANDIM: TPanel
      Left = 1
      Top = 20
      Width = 678
      Height = 20
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 0
      object TGA_CODEDIM5: TLabel
        Left = 500
        Top = 4
        Width = 32
        Height = 13
        Caption = 'Label1'
      end
      object TGA_CODEDIM4: TLabel
        Left = 376
        Top = 4
        Width = 32
        Height = 13
        Caption = 'Label1'
      end
      object TGA_CODEDIM3: TLabel
        Left = 252
        Top = 4
        Width = 32
        Height = 13
        Caption = 'Label1'
      end
      object TGA_CODEDIM2: TLabel
        Left = 132
        Top = 4
        Width = 32
        Height = 13
        Caption = 'Label1'
      end
      object TGA_CODEDIM1: TLabel
        Left = 12
        Top = 4
        Width = 32
        Height = 13
        Caption = 'Label1'
      end
    end
    object Ptotaux: TPanel
      Left = 1
      Top = 1
      Width = 678
      Height = 26
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      object Label1: TLabel
        Left = 100
        Top = 6
        Width = 60
        Height = 13
        Caption = 'Total budget'
      end
      object Label2: TLabel
        Left = 388
        Top = 6
        Width = 54
        Height = 13
        Caption = 'Total achat'
      end
      object TmontantBudget: THNumEdit
        Left = 166
        Top = 2
        Width = 85
        Height = 21
        TabOrder = 0
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TmontantAch: THNumEdit
        Left = 468
        Top = 2
        Width = 85
        Height = 21
        TabOrder = 1
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Cette suggestion n'#39'existe pas;W;O;O;O;'
      '')
    Left = 384
    Top = 144
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 100
    Top = 108
  end
  object POPZ: TPopupMenu
    Left = 48
    Top = 152
  end
  object HPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Confirmez-vous la suppression de cette suggestion de' +
        ' r'#233'approvisionnement ?;Q;YN;N;N;'
      '1;?caption?;La pi'#232'ce ne peut '#234'tre supprim'#233'e ;E;O;O;O;'
      
        '2;?caption?;Erreur ! la pi'#232'ce n'#39'a pas '#233't'#233' enregistr'#233'e ! ;E;O;O;O' +
        ';'
      
        '3;?caption?;Erreur ! vous ne pouvez pas indiquer une quantit'#233' in' +
        'f'#233'rieure au besoin mini. ! ;E;O;O;O;'
      '4;?caption?;Ce fournisseur est inconnu ou ferm'#233' ;E;O;O;O;'
      ''
      '')
    Left = 533
    Top = 167
  end
  object HTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ATTENTION : la suppression ne s'#39'est pas correctement effectu'#233'e !')
    Left = 489
    Top = 168
  end
  object PChoixFour: TPopupMenu
    AutoPopup = False
    Left = 588
    Top = 328
    object Catalogue: TMenuItem
      Caption = 'Catalogue'
      OnClick = CatalogueClick
    end
    object Listefournisseurs: TMenuItem
      Caption = 'Liste fournisseurs'
      OnClick = ListefournisseursClick
    end
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 620
    Top = 176
  end
end

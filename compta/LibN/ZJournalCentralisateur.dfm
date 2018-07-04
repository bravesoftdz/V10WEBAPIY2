object FJournalCentralisateur: TFJournalCentralisateur
  Left = 130
  Top = 153
  Width = 729
  Height = 478
  Caption = 'Journal centralisateur'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  PopupMenu = PopF11
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter: TSplitter
    Left = 285
    Top = 41
    Height = 367
    OnMoved = SplitterMoved
  end
  object FlashingLabel1: TFlashingLabel
    Left = 96
    Top = 76
    Width = 71
    Height = 13
    Caption = 'FlashingLabel1'
  end
  object DockBottom: TDock97
    Left = 0
    Top = 408
    Width = 713
    Height = 31
    AllowDrag = False
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 629
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 629
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Valider la saisie'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object Toolbar971: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BZoom: TToolbarButton97
        Left = 40
        Top = 0
        Width = 40
        Height = 26
        Alignment = taLeftJustify
        DropdownArrow = True
        DropdownMenu = POPZoom
        GlobalIndexImage = 'Z0061_S16G1'
      end
      object BNew: TToolbarButton97
        Left = 80
        Top = 0
        Width = 28
        Height = 26
        Hint = 'Nouveau'
        Alignment = taLeftJustify
        ParentShowHint = False
        ShowHint = True
        OnClick = BNewClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BAFF: TToolbarButton97
        Left = 0
        Top = 0
        Width = 40
        Height = 26
        DropdownArrow = True
        DropdownMenu = POPAFF
        GlobalIndexImage = 'M0004_S16G1'
      end
    end
  end
  object TW: TTreeView
    Left = 0
    Top = 41
    Width = 285
    Height = 367
    Align = alLeft
    ChangeDelay = 50
    HideSelection = False
    Images = imTSE
    Indent = 21
    ReadOnly = True
    RowSelect = True
    TabOrder = 1
    OnChange = TWChange
    OnClick = TWClick
    OnExpanded = TWExpanded
  end
  object FLISTE: THGrid
    Left = 288
    Top = 41
    Width = 425
    Height = 367
    Align = alClient
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
    TabOrder = 2
    OnDblClick = FLISTEDblClick
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    DBIndicator = True
    ColWidths = (
      10
      64
      64
      64
      64)
  end
  object Dock971: TDock97
    Left = 0
    Top = 0
    Width = 713
    Height = 41
    AllowDrag = False
    FixAlign = True
    object PEntete: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 37
      ClientWidth = 713
      Caption = 'Filtres'
      ClientAreaHeight = 37
      ClientAreaWidth = 713
      DockableTo = []
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        713
        37)
      object BCherche: TToolbarButton97
        Left = 675
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Anchors = [akTop, akRight]
        Flat = False
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object TEXERCICE: THLabel
        Left = 15
        Top = 9
        Width = 41
        Height = 13
        Caption = 'Exercice'
        FocusControl = EXERCICE
        Transparent = True
      end
      object TETABLISSEMENT: THLabel
        Left = 289
        Top = 9
        Width = 65
        Height = 13
        Caption = 'Etablissement'
        FocusControl = EXERCICE
        Transparent = True
      end
      object EXERCICE: THValComboBox
        Left = 63
        Top = 5
        Width = 200
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = EXERCICEChange
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object ETABLISSEMENT: THMultiValComboBox
        Left = 362
        Top = 5
        Width = 200
        Height = 21
        TabOrder = 1
        OnChange = ETABLISSEMENTChange
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
    end
  end
  object HMTrad: THSystemMenu
    Separator = True
    Traduction = False
    Left = 52
    Top = 264
  end
  object POPAFF: TPopupMenu
    Left = 160
    Top = 220
    object MMAFFPERIODE: TMenuItem
      Caption = 'Affichage par &p'#233'riode'
      ShortCut = 32848
      OnClick = MMAFFPERIODEClick
    end
    object MMAFFJAL: TMenuItem
      Caption = 'Affichage par &journal'
      ShortCut = 32842
      OnClick = MMAFFJALClick
    end
    object MMAFFNOEUDVIDE: TMenuItem
      Caption = '&Afficher les '#233'l'#233'ments non mouvement'#233's'
      ShortCut = 32846
      OnClick = MMAFFNOEUDVIDEClick
    end
  end
  object POPZoom: TPopupMenu
    Left = 160
    Top = 268
    object MMZoomJal: TMenuItem
      Caption = 'Acc'#233'der aux param'#232'tres du journal'
      ShortCut = 16458
      OnClick = MMZoomJalClick
    end
  end
  object ImXP: THImageList
    GlobalIndexImages.Strings = (
      'Z1052_S16G1'
      'Z0930_S16G1'
      'Z0930_S16G1'
      'Z1005_S16G1'
      'Z0021_S16G1'
      'Z0859_S16G1'
      'Z0369_S16G1'
      'Z0053_S16G1'
      'Z0016_S16G1')
    Left = 80
    Top = 104
  end
  object PopF11: TPopupMenu
    Images = imTSE
    OnPopup = PopF11Popup
    Left = 160
    Top = 164
    object AFFPeriode: TMenuItem
      Caption = 'Affichage par &p'#233'riode'
      ShortCut = 32848
      OnClick = AFFPeriodeClick
    end
    object AFFJournal: TMenuItem
      Caption = 'Affichage par &journal'
      ShortCut = 32842
      OnClick = AFFJournalClick
    end
    object AffNoeudVide: TMenuItem
      Caption = '&Afficher les '#233'l'#233'ments non mouvement'#233's'
      ShortCut = 32846
      OnClick = AffNoeudVideClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object NouvelElement: TMenuItem
      Caption = '&Nouveau'
      ImageIndex = 7
      ShortCut = 16462
      OnClick = NouvelElementClick
    end
    object AccesJournal: TMenuItem
      Caption = 'Acc'#233'der aux &param'#232'tres du journal'
      ShortCut = 16458
      OnClick = AccesJournalClick
    end
    object AccesDetail: TMenuItem
      Caption = 'Acc'#233'der au &d'#233'tail'
      ImageIndex = 8
      ShortCut = 116
      OnClick = AccesDetailClick
    end
    object AppliquerLesCriteres: TMenuItem
      Caption = 'Appliquer les crit'#232'res'
      ShortCut = 120
      OnClick = AppliquerLesCriteresClick
    end
    object Reclassement: TMenuItem
      Caption = '&Reclassement des donn'#233'es'
      Visible = False
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object Imprimer: TMenuItem
      Caption = '&Editions'
      ImageIndex = 6
      object JALCECR: TMenuItem
        Caption = 'Le &journal des '#233'critures'
        ShortCut = 16464
        OnClick = JALCECRClick
      end
      object JALCENTRAL: TMenuItem
        Caption = 'Le journal &centralisateur'
        OnClick = JALCENTRALClick
      end
      object BALGENE: TMenuItem
        Caption = 'La &balance g'#233'n'#233'rale'
        OnClick = BALGENEClick
      end
      object GLECR: TMenuItem
        Caption = 'Le &grand-livre g'#233'n'#233'ral'
        OnClick = GLECRClick
      end
    end
  end
  object imTSE: THImageList
    GlobalIndexImages.Strings = (
      'Z1837_S16G1'
      'Z0930_S16G1'
      'Z0930_S16G1'
      'Z1547_S16G1'
      'Z0021_S16G1'
      'Z0652_S16G1'
      'Z0369_S16G1'
      'Z0802_S16G1')
    Height = 18
    Width = 18
    Left = 105
    Top = 216
  end
end

object FVerSolde: TFVerSolde
  Left = 248
  Top = 157
  Width = 435
  Height = 375
  HelpContext = 7745300
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Solde des comptes'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 427
    Height = 85
    ActivePage = PCritere
    Align = alTop
    TabOrder = 1
    TabWidth = 85
    object PCritere: TTabSheet
      Caption = 'Crit'#232'res'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 419
        Height = 57
        Align = alClient
      end
      object TFCptes: THLabel
        Left = 9
        Top = 20
        Width = 41
        Height = 13
        Caption = '&Comptes'
        FocusControl = FCptes
      end
      object FCptes: TEdit
        Left = 85
        Top = 16
        Width = 324
        Height = 21
        TabOrder = 0
      end
    end
  end
  object PFiltres: TPanel
    Left = 0
    Top = 85
    Width = 427
    Height = 35
    Align = alTop
    Caption = ' '
    TabOrder = 2
    object BFiltre: TToolbarButton97
      Left = 6
      Top = 7
      Width = 58
      Height = 21
      Hint = 'Menu filtre'
      DropdownArrow = True
      DropdownMenu = POPF
      Caption = '&Filtres'
      Layout = blGlyphRight
      ParentShowHint = False
      ShowHint = True
    end
    object BCherche: THBitBtn
      Left = 392
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Appliquer crit'#232'res'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BChercheClick
      GlobalIndexImage = 'Z0217_S16G1'
    end
    object FFiltres: THValComboBox
      Left = 69
      Top = 7
      Width = 317
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FFiltresChange
      TagDispatch = 0
    end
  end
  object Formateur: THNumEdit
    Left = 306
    Top = 26
    Width = 85
    Height = 21
    Color = clYellow
    Decimals = 2
    Digits = 12
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Masks.PositiveMask = '#,##0'
    Debit = False
    ParentFont = False
    TabOrder = 3
    UseRounding = True
    Validate = False
    Visible = False
  end
  object FListe: THGrid
    Left = 0
    Top = 120
    Width = 427
    Height = 193
    Align = alClient
    ColCount = 3
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
    TabOrder = 4
    SortedCol = -1
    Titres.Strings = (
      'Compte'
      'Libell'#233
      'Solde;D;')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      108
      184
      113)
  end
  object PanelBouton: TPanel
    Left = 0
    Top = 313
    Width = 427
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object BReduire: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Afficher les crit'#232'res'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BReduireClick
      GlobalIndexImage = 'Z0910_S16G1'
    end
    object BAgrandir: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Agrandir la liste'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BAgrandirClick
      GlobalIndexImage = 'Z0270_S16G1'
    end
    object BRechercher: THBitBtn
      Left = 36
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Rechercher dans la liste'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BRechercherClick
      GlobalIndexImage = 'Z0077_S16G1'
    end
    object Panel1: TPanel
      Left = 293
      Top = 2
      Width = 132
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 3
      object BImprimer: THBitBtn
        Left = 35
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BFermer: THBitBtn
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: THBitBtn
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 40
    Top = 179
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 160
    Top = 179
  end
  object POPF: TPopupMenu
    OnPopup = POPFPopup
    Left = 108
    Top = 180
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      OnClick = BCreerFiltreClick
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BSaveFiltreClick
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
      OnClick = BDelFiltreClick
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
      OnClick = BRenFiltreClick
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
      OnClick = BNouvRechClick
    end
  end
end

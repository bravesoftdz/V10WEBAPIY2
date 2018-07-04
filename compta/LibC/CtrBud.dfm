object FCtrBud: TFCtrBud
  Left = 101
  Top = 140
  Width = 630
  Height = 379
  HelpContext = 15126000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Controle du param'#233'trage des comptes budg'#233'taires'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 316
    Width = 622
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 622
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        622
        32)
      object BReduire: TToolbarButton97
        Left = 4
        Top = 4
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'R'#233'duire'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 4
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Agrandir'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 492
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BOuvrir: TToolbarButton97
        Left = 524
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BOuvrirClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BAnnuler: TToolbarButton97
        Left = 556
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 588
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BRechercher: TToolbarButton97
        Left = 37
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercherClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BStop: TToolbarButton97
        Tag = 1
        Left = 70
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Arr'#234'ter le traitement'
        DisplayMode = dmGlyphOnly
        Caption = 'Arr'#234'ter'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BStopClick
        GlobalIndexImage = 'Z0107_S16G1'
        IsControl = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 103
        Top = 4
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopZ
        Caption = 'Zooms'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 622
    Height = 109
    ActivePage = ChoixEtat
    Align = alTop
    TabOrder = 1
    object ChoixEtat: TTabSheet
      Caption = 'Param'#233'trage'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 81
        Align = alClient
      end
      object TCpte: TLabel
        Left = 161
        Top = 35
        Width = 51
        Height = 13
        Caption = '&Compte de'
      end
      object Tcpt1: TLabel
        Left = 435
        Top = 35
        Width = 6
        Height = 13
        Caption = #224
      end
      object TCBud1: THLabel
        Left = 161
        Top = 10
        Width = 108
        Height = 13
        Caption = '&Comptes bug'#233'taires de'
      end
      object TCBud2: THLabel
        Left = 435
        Top = 10
        Width = 6
        Height = 13
        Caption = #224
      end
      object TCbAxe: TLabel
        Left = 161
        Top = 60
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = CbAxe
      end
      object C1: THCpteEdit
        Left = 275
        Top = 31
        Width = 150
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 35
        TabOrder = 1
        ZoomTable = tzSection
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object C2: THCpteEdit
        Left = 452
        Top = 31
        Width = 150
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 35
        TabOrder = 2
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object CbTous: TCheckBox
        Left = 435
        Top = 60
        Width = 171
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Voir tous les comptes'
        TabOrder = 3
      end
      object RgChoix: TRadioGroup
        Left = 6
        Top = 4
        Width = 142
        Height = 69
        Caption = ' Choix du contr'#244'le '
        ItemIndex = 0
        Items.Strings = (
          'Comptes &budg'#233'taires'
          'Comptes &g'#233'n'#233'raux')
        TabOrder = 0
        OnClick = RgChoixClick
      end
      object CbAxe: THValComboBox
        Left = 275
        Top = 56
        Width = 150
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnChange = CbAxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object CBud1: THCpteEdit
        Left = 275
        Top = 6
        Width = 150
        Height = 21
        TabOrder = 5
        ZoomTable = tzBudSec1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object CBud2: THCpteEdit
        Left = 452
        Top = 6
        Width = 150
        Height = 21
        TabOrder = 6
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 148
    Width = 622
    Height = 168
    Align = alClient
    ColCount = 6
    DefaultColWidth = 100
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowSelect]
    TabOrder = 0
    OnDblClick = FListeDblClick
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Intitul'#233
      'Code'
      'Intitul'#233
      'Est exclu')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      90
      210
      95
      204
      75
      75)
  end
  object BBud: THBitBtn
    Tag = 100
    Left = 285
    Top = 245
    Width = 33
    Height = 27
    Hint = 'Voir Rubrique'
    Caption = 'Bud'
    TabOrder = 2
    Visible = False
    OnClick = BBudClick
  end
  object BGen: THBitBtn
    Tag = 100
    Left = 322
    Top = 245
    Width = 33
    Height = 27
    Caption = 'Gen'
    TabOrder = 3
    Visible = False
    OnClick = BGenClick
  end
  object Dock971: TDock97
    Left = 0
    Top = 109
    Width = 622
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 622
      ClientAreaHeight = 35
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 5
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BCherche: TToolbarButton97
        Left = 586
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 68
        Top = 5
        Width = 510
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object PopZ: TPopupMenu
    Left = 55
    Top = 244
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 101
    Top = 244
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 149
    Top = 244
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Comptes &budg'#233'taires'
      'C&omptes g'#233'n'#233'raux'
      'Sections &budg'#233'taires'
      'Sections &analytiques'
      'Contr'#244'le des comptes budg'#233'taires'
      'Contr'#244'le des sections budg'#233'taires'
      '&Comptes bug'#233'taires de'
      'Co&mptes g'#233'n'#233'raux de'
      '&Sections budg'#233'taires de'
      'S&ections analytiques de'
      '10;'
      'Non'
      'Oui'
      'Mouvement'#233
      'Chantier'
      
        '15;?Caption?;Confirmez-vous l'#39'arr'#234't du traitement en cours ?;Q;Y' +
        'N;N;N;'
      'Voir compte budg'#233'taire'
      'Voir section budg'#233'taire'
      'Voir compte g'#233'n'#233'ral'
      'Voir section analytique'
      'Voir toutes les sections'
      '21;')
    Left = 195
    Top = 244
  end
  object POPF: TPopupMenu
    Left = 240
    Top = 244
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
    end
  end
end

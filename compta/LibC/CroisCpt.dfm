object FCroisCpt: TFCroisCpt
  Left = 250
  Top = 163
  Width = 636
  Height = 430
  HelpContext = 15412000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Param'#233'trage du croisement des comptes'
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
    Top = 364
    Width = 628
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 628
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 628
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 593
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
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 561
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 530
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
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
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 498
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
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
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 189
    Width = 628
    Height = 175
    Align = alClient
    ColCount = 6
    DefaultColWidth = 100
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goTabs]
    TabOrder = 1
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
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
      100
      100
      101
      100
      100
      100)
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 628
    Height = 150
    ActivePage = PCritere
    Align = alTop
    TabOrder = 2
    object PCritere: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 620
        Height = 122
        Align = alClient
      end
      object TBudJal: TLabel
        Left = 4
        Top = 9
        Width = 34
        Height = 13
        Caption = 'B&udget'
        FocusControl = BudJal
      end
      object TCbGen: TLabel
        Left = 4
        Top = 34
        Width = 114
        Height = 13
        Caption = 'C&omptes budg'#233'taires de'
        FocusControl = CbGen
      end
      object TCbGen_: TLabel
        Left = 322
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = BudJal
      end
      object TCbSec: TLabel
        Left = 4
        Top = 59
        Width = 109
        Height = 13
        Caption = 'S&ection budg'#233'taires de'
        FocusControl = CbSec
      end
      object HLabel3: THLabel
        Left = 299
        Top = 9
        Width = 59
        Height = 13
        Caption = '&Pr'#233'sentation'
      end
      object TCbSec_: TLabel
        Left = 322
        Top = 59
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = CbSec_
      end
      object Label1: TLabel
        Left = 4
        Top = 102
        Width = 470
        Height = 13
        Caption = 
          'Cochez/d'#233'cochez les restrictions budg'#233'taires par "ESPACE" ou "CT' +
          'RL+Click" sur la multi-s'#233'lection'
      end
      object TypeCroisement: TRadioGroup
        Left = 533
        Top = 4
        Width = 78
        Height = 72
        Caption = ' Croiser sur '
        ItemIndex = 0
        Items.Strings = (
          '&Budget'
          '&Cat'#233'gorie')
        TabOrder = 0
        OnClick = TypeCroisementClick
      end
      object BudJal: THValComboBox
        Left = 125
        Top = 5
        Width = 160
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = BudJalChange
        TagDispatch = 0
        DataType = 'TTBUDJALSANSCAT'
      end
      object CBGenere: TCheckBox
        Left = 2
        Top = 79
        Width = 521
        Height = 19
        Alignment = taLeftJustify
        Caption = 
          '&G'#233'n'#233'rer les croisements pour le contr'#244'le budg'#233'taire en saisie d' +
          'e comptabilit'#233' g'#233'n'#233'rale'
        TabOrder = 2
      end
      object CbGen: THValComboBox
        Left = 125
        Top = 30
        Width = 160
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        Vide = True
      end
      object TypVue: THValComboBox
        Left = 364
        Top = 5
        Width = 160
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnChange = TypVueChange
        Items.Strings = (
          'Comptes / Section'
          'Sections / Compte')
        TagDispatch = 0
        Values.Strings = (
          'GS'
          'SG')
      end
      object CbGen_: THValComboBox
        Left = 364
        Top = 30
        Width = 160
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        TagDispatch = 0
        Vide = True
      end
      object CbSec: THValComboBox
        Left = 125
        Top = 55
        Width = 160
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        TagDispatch = 0
        Vide = True
      end
      object CbSec_: THValComboBox
        Left = 364
        Top = 55
        Width = 160
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        TagDispatch = 0
        Vide = True
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 150
    Width = 628
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 628
      ClientAreaHeight = 35
      ClientAreaWidth = 628
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
        Left = 587
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 69
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
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 536
    Top = 256
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Param'#233'trage du croisement des comptes;D'#233'sirez-vous valider vot' +
        're nouvelle s'#233'lection?;Q;YN;N;N;'
      '&Budget'
      '&Cat'#233'gorie'
      
        '3;Param'#233'trage du croisement des comptes;La fourchette des compte' +
        's d'#233'finie est incoh'#233'rente;W;O;O;O;'
      'Gestion des croisements')
    Left = 500
    Top = 256
  end
  object POPF: TPopupMenu
    Left = 152
    Top = 276
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

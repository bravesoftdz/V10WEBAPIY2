object FCodBuRup: TFCodBuRup
  Left = 185
  Top = 212
  Width = 621
  Height = 370
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Cr'#233'ation de comptes '#224' partir d'#39'un plan de rupture'
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 613
    Height = 115
    ActivePage = Pparam
    Align = alTop
    TabOrder = 0
    object Pparam: TTabSheet
      Caption = 'Param'#233'trage'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 605
        Height = 87
        Align = alClient
      end
      object TSens: THLabel
        Left = 303
        Top = 18
        Width = 24
        Height = 13
        Caption = '&Sens'
        FocusControl = Sens
      end
      object TAxe: THLabel
        Left = 23
        Top = 54
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Axe
      end
      object TCbRupt: TLabel
        Left = 23
        Top = 18
        Width = 72
        Height = 13
        Caption = '&Plan de rupture'
      end
      object Nb1: TLabel
        Left = 365
        Top = 54
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Tex1: TLabel
        Left = 390
        Top = 54
        Width = 132
        Height = 13
        Caption = 'ligne(s) s'#233'lectionn'#233'e(s)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object BCherche: TToolbarButton97
        Left = 562
        Top = 11
        Width = 28
        Height = 27
        Hint = 'Rechercher les combinaisons'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object Sens: THValComboBox
        Left = 348
        Top = 14
        Width = 189
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTSENS'
      end
      object Signe: THValComboBox
        Left = 300
        Top = 46
        Width = 29
        Height = 21
        Style = csDropDownList
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 3
        Visible = False
        TagDispatch = 0
        DataType = 'TTRUBSIGNE'
      end
      object Axe: THValComboBox
        Left = 100
        Top = 50
        Width = 189
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = AxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object CbRupt: THValComboBox
        Left = 100
        Top = 14
        Width = 189
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTRUPTGENE'
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 115
    Width = 613
    Height = 193
    Align = alClient
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    TabOrder = 1
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Libell'#233
      'Rubrique'
      'Code rubrique')
    Couleur = True
    MultiSelect = True
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      120
      195
      150
      125
      0)
  end
  object Dock: TDock97
    Left = 0
    Top = 308
    Width = 613
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 609
      Caption = 'Actions'
      ClientAreaHeight = 31
      ClientAreaWidth = 609
      DockPos = 0
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      DesignSize = (
        609
        31)
      object BTag: TToolbarButton97
        Left = 3
        Top = 3
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'S'#233'lectionner'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BTagClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 484
        Top = 3
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
      object BValider: TToolbarButton97
        Left = 516
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Lancer la cr'#233'ation'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Cr'#233'er'
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
      object BFerme: TToolbarButton97
        Left = 548
        Top = 3
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
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 580
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
      object Bdetag: TToolbarButton97
        Left = 3
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'selectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'selectionner'
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BdetagClick
        GlobalIndexImage = 'Z0078_S16G2'
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 180
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Comptes g'#233'n'#233'raux'
      'Sections analytiques'
      'Ligne s'#233'lectionn'#233'e'
      'Lignes s'#233'lectionn'#233'es'
      'G'#233'n'#233'ration des comptes budg'#233'taires '#224' partir d'#39'un plan de rupture'
      
        'G'#233'n'#233'ration des sections budg'#233'taires '#224' partir d'#39'un plan de ruptur' +
        'e'
      
        '6;G'#233'n'#233'ration des codes budg'#233'taires; Vous n'#39'avez rien s'#233'lectionn'#233 +
        '. Aucun traitement '#224' effectuer.;W;O;O;O;'
      
        '7;G'#233'n'#233'ration des codes budg'#233'taires;D'#233'sirez vous cr'#233'er les codes ' +
        'budg'#233'taires de votre s'#233'lection?;Q;YNC;N;N;'
      
        '8;G'#233'n'#233'ration des codes budg'#233'taires;Des codes budg'#233'taires n'#39'ont p' +
        'as '#233't'#233' cr'#233#233's. D'#233'sirez-vous les modifier pour les cr'#233#233'r;Q;YNC;Y;N' +
        ';'
      'ATTENTION : Cr'#233'ation non effectu'#233'e !'
      
        '10;G'#233'n'#233'ration des codes budg'#233'taires;La cr'#233'ation des codes s'#39'est ' +
        'correctement effectu'#233'e !;I;O;O;O;')
    Left = 408
    Top = 196
  end
end

object CorrespDlg: TCorrespDlg
  Left = 315
  Top = 247
  HelpContext = 1210
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Table de correspondance'
  ClientHeight = 293
  ClientWidth = 626
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
  object Label1: TLabel
    Left = 422
    Top = 100
    Width = 97
    Height = 13
    Caption = 'Valeurs d'#233'j'#224' d'#233'finies'
  end
  object Label2: TLabel
    Left = 80
    Top = 100
    Width = 128
    Height = 13
    Caption = 'Nouvelles valeurs trouv'#233'es'
  end
  object Compteur: TLabel
    Left = 8
    Top = 260
    Width = 3
    Height = 13
  end
  object Label16: TLabel
    Left = 7
    Top = 67
    Width = 69
    Height = 13
    Caption = 'Fichier externe'
  end
  object Label3: TLabel
    Left = 246
    Top = 6
    Width = 30
    Height = 13
    Caption = 'Libell'#233
  end
  object Label8: TLabel
    Left = 7
    Top = 35
    Width = 50
    Height = 13
    Caption = 'Code profil'
  end
  object Label12: TLabel
    Left = 306
    Top = 34
    Width = 32
    Height = 13
    Caption = 'Famille'
  end
  object Breprise: TToolbarButton97
    Left = 591
    Top = 31
    Width = 33
    Height = 20
    Hint = 'Selection par famille'
    HelpContext = 1710
    DropdownArrow = True
    DropdownMenu = POPZ
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Margin = 3
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    OnMouseDown = FamilleMouseDown
    GlobalIndexImage = 'Z1334_S16G1'
    IsControl = True
  end
  object TDOMAINE: TLabel
    Left = 7
    Top = 6
    Width = 42
    Height = 13
    Caption = 'Domaine'
  end
  object Button1: TButton
    Left = 268
    Top = 124
    Width = 75
    Height = 25
    Caption = '>'
    TabOrder = 5
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 268
    Top = 156
    Width = 75
    Height = 25
    Caption = '<'
    TabOrder = 6
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 268
    Top = 188
    Width = 75
    Height = 25
    Caption = '>>'
    TabOrder = 8
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 268
    Top = 220
    Width = 75
    Height = 25
    Caption = '<<'
    TabOrder = 9
    OnClick = Button4Click
  end
  object cbOpTblCorr: TComboBox
    Left = 400
    Top = 66
    Width = 87
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 7
    Visible = False
    OnChange = cbOpTblCorrChange
    Items.Strings = (
      'Toujours parcourir sans confirmation'
      'Parcourir seulement sur confirmation'
      'Jamais parcourir'
      'D'#233'faut')
  end
  object Dock971: TDock97
    Left = 0
    Top = 258
    Width = 626
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 626
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 31
      ClientAreaWidth = 626
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        626
        31)
      object btnOK: TToolbarButton97
        Left = 476
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = btnOKClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object btnFermer: TToolbarButton97
        Left = 508
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Cancel = True
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object bDefaire: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Annuler les modifications'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object Binsert: TToolbarButton97
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        AllowAllUp = True
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
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 442
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        AllowAllUp = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnSupprimerClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 116
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
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
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Tag = 1210
        Left = 587
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = HelpBtnClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object btnScruter: THBitBtn
        Left = 409
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Scruter'
        HelpContext = 1710
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = btnScruterClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0016_S16G1'
        IsControl = True
      end
    end
  end
  object edNomTable: THCritMaskEdit
    Left = 80
    Top = 63
    Width = 508
    Height = 21
    AutoSelect = False
    TabOrder = 3
    TagDispatch = 0
    DataType = 'OPENFILE(*.TXT;*.*)'
    ElipsisButton = True
  end
  object StringGrid3: TStringGrid
    Left = 28
    Top = 115
    Width = 225
    Height = 17
    BorderStyle = bsNone
    ColCount = 3
    DefaultColWidth = 115
    DefaultRowHeight = 30
    FixedCols = 2
    RowCount = 1
    FixedRows = 0
    TabOrder = 11
  end
  object StringGrid4: TStringGrid
    Left = 356
    Top = 115
    Width = 241
    Height = 17
    BorderStyle = bsNone
    ColCount = 3
    DefaultColWidth = 120
    DefaultRowHeight = 30
    FixedCols = 2
    RowCount = 1
    FixedRows = 0
    TabOrder = 12
    OnDblClick = StringGrid4DblClick
  end
  object Profile: THCritMaskEdit
    Left = 80
    Top = 31
    Width = 208
    Height = 21
    MaxLength = 20
    TabOrder = 1
    TagDispatch = 0
  end
  object Famille: THCritMaskEdit
    Left = 345
    Top = 31
    Width = 240
    Height = 21
    MaxLength = 35
    TabOrder = 2
    OnMouseDown = FamilleMouseDown
    TagDispatch = 0
  end
  object Memo1: THCritMaskEdit
    Left = 286
    Top = 2
    Width = 329
    Height = 21
    MaxLength = 35
    TabOrder = 0
    TagDispatch = 0
  end
  object Domaine: THValComboBox
    Left = 80
    Top = 2
    Width = 125
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 13
    TagDispatch = 0
    ComboWidth = 300
  end
  object StringGrid2: THGrid
    Left = 26
    Top = 137
    Width = 228
    Height = 120
    ColCount = 2
    DefaultColWidth = 125
    DefaultRowHeight = 17
    FixedCols = 0
    RowCount = 2
    FixedRows = 0
    ScrollBars = ssVertical
    TabOrder = 4
    OnEnter = StringGrid1Enter
    OnKeyDown = StringGrid2KeyDown
    OnSelectCell = SelectCell
    OnSetEditText = StringGrid1SetEditText
    OnTopLeftChanged = StringGrid2TopLeftChanged
    SortedCol = 0
    Couleur = True
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
  end
  object StringGrid1: THGrid
    Left = 354
    Top = 137
    Width = 243
    Height = 120
    ColCount = 2
    DefaultColWidth = 125
    DefaultRowHeight = 16
    FixedCols = 0
    RowCount = 2
    FixedRows = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing, goThumbTracking]
    ScrollBars = ssVertical
    TabOrder = 14
    OnClick = StringGrid1Click
    OnEnter = StringGrid1Enter
    OnKeyDown = StringGrid1KeyDown
    OnSelectCell = SelectCell
    OnSetEditText = StringGrid1SetEditText
    SortedCol = 0
    Couleur = True
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 176
    Top = 156
  end
  object POPZ: TPopupMenu
    Left = 505
    Top = 161
  end
end

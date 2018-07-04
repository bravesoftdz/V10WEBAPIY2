object FEdtLegal: TFEdtLegal
  Left = 264
  Top = 149
  Width = 595
  Height = 328
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Etat des '#233'ditions'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 587
    Height = 85
    ActivePage = PCritere
    Align = alTop
    TabOrder = 0
    TabWidth = 85
    object PCritere: TTabSheet
      Caption = 'Crit'#232'res'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 579
        Height = 57
        Align = alClient
      end
      object TED_TYPEEDITION: THLabel
        Left = 9
        Top = 10
        Width = 32
        Height = 13
        Caption = '&Edition'
        FocusControl = ED_TYPEEDITION
      end
      object TED_EXERCICE: THLabel
        Left = 275
        Top = 10
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = ED_EXERCICE
      end
      object TED_UTILISATEUR: THLabel
        Left = 9
        Top = 34
        Width = 46
        Height = 13
        Caption = '&Utilisateur'
        FocusControl = ED_UTILISATEUR
      end
      object TED_DATEEDITION: THLabel
        Left = 275
        Top = 34
        Width = 77
        Height = 13
        Caption = '&Dates '#233'dition du'
        FocusControl = ED_DATEEDITION
      end
      object TTED_DATEEDITION: THLabel
        Left = 460
        Top = 34
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = ED_DATEEDITION_
      end
      object ED_TYPEEDITION: THValComboBox
        Left = 71
        Top = 6
        Width = 185
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTEDITIONLEGALE'
      end
      object ED_EXERCICE: THValComboBox
        Left = 364
        Top = 6
        Width = 201
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object XX_WHERE: TEdit
        Left = 256
        Top = 20
        Width = 29
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Text = 'ED_TYPEEDITION<>"TIS"'
        Visible = False
      end
      object ED_UTILISATEUR: THValComboBox
        Left = 71
        Top = 30
        Width = 185
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTUTILISATEUR'
      end
      object ED_DATEEDITION: THCritMaskEdit
        Left = 364
        Top = 30
        Width = 80
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 4
        Text = '01/01/1900'
        OnKeyPress = ED_DATEEDITIONKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object ED_DATEEDITION_: THCritMaskEdit
        Left = 485
        Top = 30
        Width = 80
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 5
        Text = '31/12/2099'
        OnKeyPress = ED_DATEEDITIONKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object ED_OBLIGATOIRE: TCheckBox
        Left = 332
        Top = 4
        Width = 97
        Height = 17
        AllowGrayed = True
        Caption = 'ED_OBLIGATOIRE'
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 6
        Visible = False
      end
    end
  end
  object ECaption: TEdit
    Left = 284
    Top = 152
    Width = 153
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    Text = 'l'#233'gales'
    Visible = False
  end
  object ECaption2: TEdit
    Left = 284
    Top = 176
    Width = 153
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    Text = 'R'#233'capitulatif des '#233'ditions'
    Visible = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 85
    Width = 587
    Height = 37
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 33
      ClientWidth = 587
      Caption = ' '
      ClientAreaHeight = 33
      ClientAreaWidth = 587
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Flat = False
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BCherche: TToolbarButton97
        Left = 548
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 68
        Top = 7
        Width = 477
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 253
    Width = 587
    Height = 41
    AllowDrag = False
    Position = dpBottom
    object PanelBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 37
      ClientWidth = 587
      Caption = 'PanelBouton'
      ClientAreaHeight = 37
      ClientAreaWidth = 587
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Afficher les crit'#232'res'
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
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
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
      object BParamListe: TToolbarButton97
        Left = 68
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Param'#233'trer la liste'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BParamListeClick
        GlobalIndexImage = 'Z0025_S16G1'
      end
      object BRechercher: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercherClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BPurge: TToolbarButton97
        Left = 100
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Purger'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPurgeClick
        GlobalIndexImage = 'Z0204_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 488
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
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
      object BFermer: TToolbarButton97
        Left = 520
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 552
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object FLISTE: THDBGrid
    Left = 0
    Top = 122
    Width = 587
    Height = 131
    Align = alClient
    DataSource = SEdtl
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgCancelOnExit]
    TabOrder = 5
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
  end
  object SEdtl: TDataSource
    DataSet = QEdtl
    OnDataChange = SEdtlDataChange
    Left = 120
    Top = 128
  end
  object QEdtl: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    AfterOpen = QEdtlAfterOpen
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    PageCriteres = Pages
    MAJAuto = False
    Distinct = False
    Left = 184
    Top = 128
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 40
    Top = 204
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 160
    Top = 204
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Purge du r'#233'capitulatif des '#233'ditions')
    Left = 220
    Top = 204
  end
  object POPF: TPopupMenu
    Left = 276
    Top = 204
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

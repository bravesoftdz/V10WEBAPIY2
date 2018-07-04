inherited FDPTableauBordGeneral: TFDPTableauBordGeneral
  Left = 252
  Top = 183
  Width = 821
  Height = 544
  Caption = ''
  OldCreateOrder = True
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 469
    Width = 813
    Height = 41
    inherited PBouton: TToolWindow97
      ClientHeight = 37
      ClientWidth = 813
      ClientAreaHeight = 37
      ClientAreaWidth = 813
      inherited BValider: TToolbarButton97
        Left = 401
        Visible = False
      end
      inherited BFerme: TToolbarButton97
        Left = 433
      end
      inherited HelpBtn: TToolbarButton97
        Left = 465
      end
      inherited BImprimer: TToolbarButton97
        Left = 369
        Visible = True
        OnClick = BImprimerClick
      end
      object BEffacerPaie: TToolbarButton97
        Left = 292
        Top = 2
        Width = 27
        Height = 27
        Hint = 'Effacement des donn'#233'es Paie'
        Flat = False
        OnClick = PurgerPaieClick
        GlobalIndexImage = 'Z0080_S16G1'
      end
    end
  end
  object Panel2: TPanel [1]
    Left = 0
    Top = 0
    Width = 813
    Height = 33
    Align = alTop
    Alignment = taLeftJustify
    Caption = ' Comptabilit'#233
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
  end
  object GrilleCompta: THGrid [2]
    Left = 0
    Top = 33
    Width = 813
    Height = 64
    Align = alTop
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    ScrollBars = ssNone
    TabOrder = 2
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    RowHeights = (
      18
      18
      18
      18
      18)
  end
  object Panel3: TPanel [3]
    Left = 0
    Top = 97
    Width = 813
    Height = 33
    Align = alTop
    Alignment = taLeftJustify
    Caption = ' Paie'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    object HLabel2: THLabel
      Left = 588
      Top = 10
      Width = 31
      Height = 13
      Caption = 'Ann'#233'e'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object ListeAnnee: THValComboBox
      Left = 640
      Top = 6
      Width = 120
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
      OnChange = ListeAnneeChange
      TagDispatch = 0
    end
  end
  object GrillePaie: THGrid [4]
    Left = 0
    Top = 130
    Width = 813
    Height = 134
    Align = alTop
    DefaultColWidth = 45
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    ScrollBars = ssHorizontal
    TabOrder = 4
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    RowHeights = (
      18
      18
      18
      18
      18)
  end
  object Panel4: TPanel [5]
    Left = 0
    Top = 264
    Width = 813
    Height = 33
    Align = alTop
    Alignment = taLeftJustify
    Caption = ' Gestion Interne'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 5
    object RB_MISSIONSTOUTES: THRadiobutton
      Left = 344
      Top = 8
      Width = 153
      Height = 17
      Caption = 'Toutes les missions'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = RB_MISSIONSTOUTESClick
    end
    object RB_MISSIONSENCOURS: THRadiobutton
      Left = 112
      Top = 8
      Width = 225
      Height = 17
      Caption = 'Uniquement les missions en cours'
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      TabStop = True
      OnClick = RB_MISSIONSENCOURSClick
    end
  end
  object GrilleGI: THGrid [6]
    Left = 0
    Top = 297
    Width = 813
    Height = 139
    Align = alClient
    DefaultColWidth = 80
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    PopupMenu = MenuOption
    TabOrder = 6
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    RowHeights = (
      18
      18
      18
      18
      18)
  end
  object Panel1: TPanel [7]
    Left = 0
    Top = 436
    Width = 813
    Height = 33
    Align = alBottom
    Alignment = taLeftJustify
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 7
    OnMouseMove = Panel1MouseMove
    object Solde: THLabel
      Left = 93
      Top = 10
      Width = 27
      Height = 13
      Caption = 'Solde'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object LSolde: THLabel
      Left = 6
      Top = 10
      Width = 69
      Height = 13
      Caption = 'Solde Client'
      OnClick = LSoldeClick
      OnMouseMove = LSoldeMouseMove
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 492
    Top = 124
  end
  object MenuOption: TPopupMenu
    Left = 584
    Top = 124
    object CumulAssistant: TMenuItem
      Caption = 'Cumul heures assistants'
      OnClick = CumulAssistantClick
    end
    object Dtailralise1: TMenuItem
      Caption = 'D'#233'tail r'#233'alis'#233
      OnClick = DetailRealiseClick
    end
    object Dtailfactur1: TMenuItem
      Caption = 'D'#233'tail factur'#233
      OnClick = DetailFactureClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object mnGrandLivre: TMenuItem
      Caption = 'Grand livre'
      OnClick = mnGrandLivreClick
    end
  end
end

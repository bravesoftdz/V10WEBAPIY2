object FplanningBTP: TFplanningBTP
  Left = 247
  Top = 196
  Width = 1136
  Height = 805
  Caption = 'Planning'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PCENTRE: TPanel
    Left = 0
    Top = 69
    Width = 1120
    Height = 550
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object PCENTER: TPanel
      Left = 0
      Top = 0
      Width = 1120
      Height = 550
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 0
      object G2: THGrid
        Left = 739
        Top = 0
        Width = 384
        Height = 537
        ColCount = 16
        DefaultColWidth = 35
        DefaultRowHeight = 18
        FixedCols = 0
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
        ScrollBars = ssHorizontal
        TabOrder = 0
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
      object G1: THGrid
        Left = 0
        Top = 1
        Width = 448
        Height = 537
        DefaultRowHeight = 18
        FixedCols = 0
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        ScrollBars = ssVertical
        TabOrder = 1
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
    end
  end
  object PHAUT: TPanel
    Left = 0
    Top = 0
    Width = 1120
    Height = 69
    Align = alTop
    BevelOuter = bvNone
    Locked = True
    TabOrder = 1
    object TLIBAFFAIRE: THLabel
      Left = 22
      Top = 48
      Width = 30
      Height = 13
      Caption = 'Affaire'
    end
    object TLIBELLE: THLabel
      Left = 88
      Top = 48
      Width = 118
      Height = 13
      Caption = 'R'#233'f'#233'rence && D'#233'signation'
    end
    object TLIBHRSPREV: THLabel
      Left = 264
      Top = 38
      Width = 38
      Height = 26
      Alignment = taCenter
      Caption = 'Heures pr'#233'vues'
      WordWrap = True
    end
    object TLIBRESTEAPLANNIF: THLabel
      Left = 336
      Top = 38
      Width = 40
      Height = 26
      Alignment = taCenter
      Caption = 'Reste '#224' planifier'
      WordWrap = True
    end
    object GMois: THGrid
      Left = 585
      Top = 14
      Width = 320
      Height = 19
      BorderStyle = bsNone
      ColCount = 15
      DefaultRowHeight = 18
      DefaultDrawing = False
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      ScrollBars = ssNone
      TabOrder = 0
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
    object GSem: THGrid
      Left = 585
      Top = 32
      Width = 320
      Height = 19
      BorderStyle = bsNone
      ColCount = 15
      DefaultRowHeight = 18
      DefaultDrawing = False
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      ScrollBars = ssNone
      TabOrder = 1
      OnDrawCell = GSemDrawCell
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
    object GDate: THGrid
      Left = 585
      Top = 50
      Width = 320
      Height = 19
      BorderStyle = bsNone
      ColCount = 15
      DefaultRowHeight = 18
      DefaultDrawing = False
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      ScrollBars = ssNone
      TabOrder = 2
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
  end
  object PBAS: TPanel
    Left = 0
    Top = 619
    Width = 1120
    Height = 106
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
    object TLIBTOTALHRS: THLabel
      Left = 464
      Top = 7
      Width = 79
      Height = 13
      Caption = 'Total des heures'
    end
    object TLibEffectif: THLabel
      Left = 464
      Top = 24
      Width = 38
      Height = 13
      Caption = 'Effectifs'
    end
    object LNBRRESS: TLabel
      Left = 9
      Top = 8
      Width = 155
      Height = 13
      Caption = 'Nombre de personnes '#224' affecter '
    end
    object LNBHRS: TLabel
      Left = 9
      Top = 28
      Width = 128
      Height = 13
      Caption = 'Nombre d'#39'heures '#224' affecter'
    end
    object LPOTENTIEL: THLabel
      Left = 464
      Top = 42
      Width = 41
      Height = 13
      Caption = 'Potentiel'
    end
    object LAFFECTE: THLabel
      Left = 464
      Top = 61
      Width = 34
      Height = 13
      Caption = 'Affect'#233
    end
    object LRESTE: THLabel
      Left = 464
      Top = 81
      Width = 28
      Height = 13
      Caption = 'Reste'
    end
    object LFONCTION: THLabel
      Left = 4
      Top = 60
      Width = 443
      Height = 16
      Alignment = taRightJustify
      AutoSize = False
      Caption = 'XXXXXXXXXXX'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object GTHeures: THGrid
      Left = 585
      Top = 3
      Width = 320
      Height = 21
      BorderStyle = bsNone
      ColCount = 15
      DefaultRowHeight = 18
      DefaultDrawing = False
      Enabled = False
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
      ScrollBars = ssNone
      TabOrder = 0
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
    object GTEffectifs: THGrid
      Left = 585
      Top = 21
      Width = 320
      Height = 19
      BorderStyle = bsNone
      ColCount = 15
      DefaultRowHeight = 18
      DefaultDrawing = False
      Enabled = False
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
      ScrollBars = ssNone
      TabOrder = 1
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
    object NBRRESS: THSpinEdit
      Left = 169
      Top = 4
      Width = 41
      Height = 22
      MaxValue = 99
      MinValue = 1
      TabOrder = 2
      Value = 1
      OnChange = NBRRESSChange
    end
    object NBHRS: THNumEdit
      Left = 168
      Top = 24
      Width = 85
      Height = 21
      TabOrder = 3
      OnExit = NBHRSExit
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object GFONCTION: THGrid
      Left = 585
      Top = 42
      Width = 320
      Height = 55
      BorderStyle = bsNone
      ColCount = 15
      DefaultRowHeight = 54
      DefaultDrawing = False
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
      ParentShowHint = False
      ScrollBars = ssNone
      ShowHint = True
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
  end
  object PACTIONS: TPanel
    Left = 0
    Top = 725
    Width = 1120
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 3
    DesignSize = (
      1120
      41)
    object BValider: TToolbarButton97
      Tag = 1
      Left = 1058
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Enregistrer la pi'#232'ce'
      Caption = 'Enregistrer'
      Anchors = [akRight, akBottom]
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BValiderClick
      GlobalIndexImage = 'M0012_S16G1'
      IsControl = True
    end
    object BAbandon: TToolbarButton97
      Tag = 1
      Left = 1086
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Caption = 'Annuler'
      Anchors = [akTop, akRight]
      Cancel = True
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ModalResult = 2
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BRefresh: TToolbarButton97
      Left = 1031
      Top = 3
      Width = 27
      Height = 27
      Hint = 'Raffraichir les donn'#233'es'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      OnClick = BRefreshClick
      GlobalIndexImage = 'M0063_S16G1'
    end
    object BOuvrirFermer: TToolbarButton97
      Left = 176
      Top = 8
      Width = 27
      Height = 27
      DisplayMode = dmGlyphOnly
      Images = ImgListPlusMoins
      Opaque = False
      OnClick = BOuvrirFermerClick
    end
    object BMODIF: TToolbarButton97
      Left = 1001
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Modifier la plannification'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      OnClick = BMODIFClick
      GlobalIndexImage = 'O0066_S24G1'
    end
    object BANNULE: TToolbarButton97
      Tag = 1
      Left = 1086
      Top = 3
      Width = 28
      Height = 27
      Anchors = [akRight, akBottom]
      Visible = False
      OnClick = BANNULEClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BPARAMS: TToolbarButton97
      Left = 16
      Top = 3
      Width = 27
      Height = 27
      Hint = 'Param'#232'tres'
      AllowAllUp = True
      GroupIndex = 1
      ParentShowHint = False
      ShowHint = True
      OnClick = BPARAMSClick
      GlobalIndexImage = 'O0008_S24G1'
    end
    object TBVIEWARBO: TToolbarButton97
      Left = 45
      Top = 3
      Width = 27
      Height = 27
      Hint = 'Param'#232'tres'
      AllowAllUp = True
      GroupIndex = 2
      ParentShowHint = False
      ShowHint = True
      Visible = False
      OnClick = TBVIEWARBOClick
      GlobalIndexImage = 'Z0189_S16G0'
    end
    object BImprimer: TToolbarButton97
      Left = 971
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Caption = 'Imprimer'
      Anchors = [akTop, akRight]
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
      OnClick = BImprimerClick
      GlobalIndexImage = 'Z0369_S16G1'
    end
  end
  object TTPARAMS: TToolWindow97
    Left = 440
    Top = 240
    ClientHeight = 172
    ClientWidth = 257
    Caption = 'Param'#232'tres'
    ClientAreaHeight = 172
    ClientAreaWidth = 257
    Resizable = False
    TabOrder = 4
    Visible = False
    OnClose = TTPARAMSClose
    object HLabel1: THLabel
      Left = 28
      Top = 24
      Width = 56
      Height = 13
      Caption = 'Date d'#233'part'
    end
    object HLabel2: THLabel
      Left = 28
      Top = 46
      Width = 40
      Height = 13
      Caption = 'Date Fin'
    end
    object BREFRESHP: TToolbarButton97
      Left = 213
      Top = 27
      Width = 27
      Height = 27
      Hint = 'Appliquer les dates'
      ParentShowHint = False
      ShowHint = True
      OnClick = BREFRESHPClick
      GlobalIndexImage = 'M0063_S16G1'
    end
    object DateD: THCritMaskEdit
      Left = 91
      Top = 21
      Width = 110
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 0
      Text = '  /  /    '
      TagDispatch = 0
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = True
    end
    object DateD_: THCritMaskEdit
      Left = 91
      Top = 43
      Width = 110
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '  /  /    '
      TagDispatch = 0
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = True
    end
    object CAffHeures: TCheckBox
      Left = 36
      Top = 76
      Width = 173
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Affichage des heures/jours'
      Checked = True
      State = cbChecked
      TabOrder = 2
      OnClick = CAffHeuresClick
    end
    object GroupBox1: TGroupBox
      Left = 26
      Top = 99
      Width = 197
      Height = 47
      Caption = 'Mode affichage'
      TabOrder = 3
      object RBHeures: TRadioButton
        Left = 25
        Top = 19
        Width = 59
        Height = 17
        Caption = 'Heures'
        Checked = True
        TabOrder = 0
        TabStop = True
      end
      object RbJours: TRadioButton
        Left = 120
        Top = 19
        Width = 63
        Height = 17
        Caption = 'Jours'
        TabOrder = 1
      end
    end
  end
  object TTARBO: TToolWindow97
    Left = 56
    Top = 112
    ClientHeight = 418
    ClientWidth = 305
    Caption = 'Arborescence'
    ClientAreaHeight = 418
    ClientAreaWidth = 305
    TabOrder = 5
    Visible = False
    OnClose = TTARBOClose
    object TTVARBO: TTreeView
      Left = 0
      Top = 0
      Width = 305
      Height = 418
      Align = alClient
      Indent = 19
      RowSelect = True
      TabOrder = 0
      OnClick = TTVARBOClick
    end
  end
  object ImgListPlusMoins: THImageList
    GlobalIndexImages.Strings = (
      'Z1160_S16G1'
      'Z1135_S16G1')
    Height = 18
    Width = 18
    Left = 688
    Top = 152
  end
  object POPACTIONS: TPopupMenu
    Left = 536
    Top = 103
    object PlannifCh: TMenuItem
      Caption = 'Planifier le chantier'
      OnClick = PlannifChClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object AffectRessources: TMenuItem
      Caption = 'Affecter les ressources'
      OnClick = AffectRessourcesClick
    end
  end
end

object FValJal: TFValJal
  Left = 305
  Top = 256
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'FValJal'
  ClientHeight = 346
  ClientWidth = 517
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
  object Dock: TDock97
    Left = 0
    Top = 310
    Width = 517
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 517
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 517
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object Bdetag: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'selectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233's'#233'lection'
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BdetagClick
        GlobalIndexImage = 'Z0078_S16G2'
      end
      object BTag: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'S'#233'lection'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BTagClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object BFirst: TToolbarButton97
        Left = 317
        Top = 3
        Width = 45
        Height = 27
        Hint = 'Premier'
        DisplayMode = dmGlyphOnly
        Caption = 'Premier'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFirstClick
        GlobalIndexImage = 'Z0301_S16G1'
      end
      object BPrev: TToolbarButton97
        Left = 366
        Top = 3
        Width = 45
        Height = 27
        Hint = 'Pr'#233'c'#233'dent'
        DisplayMode = dmGlyphOnly
        Caption = 'Pr'#233'c'#233'dent'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPrevClick
        GlobalIndexImage = 'Z0017_S16G1'
      end
      object BNext: TToolbarButton97
        Left = 415
        Top = 3
        Width = 45
        Height = 27
        Hint = 'Suivant'
        DisplayMode = dmGlyphOnly
        Caption = 'Suivant'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNextClick
        GlobalIndexImage = 'Z0031_S16G1'
      end
      object BLast: TToolbarButton97
        Left = 464
        Top = 3
        Width = 45
        Height = 27
        Hint = 'Dernier'
        DisplayMode = dmGlyphOnly
        Caption = 'Dernier'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BLastClick
        GlobalIndexImage = 'Z0264_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 177
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
      object BValider: TToolbarButton97
        Left = 209
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
      object BFerme: TToolbarButton97
        Left = 241
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
      object BAide: TToolbarButton97
        Left = 273
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
    end
  end
  object PExo: TPanel
    Left = 0
    Top = 0
    Width = 517
    Height = 57
    Align = alTop
    BevelInner = bvLowered
    TabOrder = 0
    object TExo: THLabel
      Left = 17
      Top = 12
      Width = 41
      Height = 13
      Caption = '&Exercice'
      FocusControl = CbExo
    end
    object BCherche: TToolbarButton97
      Left = 468
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Appliquer crit'#232'res'
      ParentShowHint = False
      ShowHint = True
      OnClick = BChercheClick
      GlobalIndexImage = 'Z0217_S16G1'
    end
    object Nb1: TLabel
      Left = 16
      Top = 35
      Width = 8
      Height = 13
      Caption = '0'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Tex1: TLabel
      Left = 36
      Top = 35
      Width = 104
      Height = 13
      Caption = 'ligne s'#233'lectionn'#233'e'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object iCritGlyph: TImage
      Left = 375
      Top = 33
      Width = 32
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D617076010000424D760100000000000076000000280000002000
        000010000000010004000000000000010000120B0000120B0000100000000000
        000000000000000080000080000000808000800000008000800080800000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00555555555555555555555555555555555555555555555555555555FF5555
        555555555C05555555555555588FF55555555555CCC05555555555558888F555
        55555555CCC05555555555558888FF555555555CCCCC05555555555888888F55
        555555CCCCCC05555555558888888FF5555554CC05CCC05555555888858888F5
        55554C05555CC05555558885555888FF55555555555CCC05555555555558888F
        555555555555CC05555555555555888FF555555555555CC05555555555555888
        FF5555555555554C05555555555555888FF5555555555554C055555555555558
        88FF5555555555555CC055555555555558885555555555555555555555555555
        5555}
      Visible = False
    end
    object iCritGlyphModified: TImage
      Left = 416
      Top = 33
      Width = 32
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D617076010000424D760100000000000076000000280000002000
        000010000000010004000000000000010000120B0000120B0000100000000000
        0000000000000000800000800000008080008000000080008000808000008080
        8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00555555555555555555555555555555555555555555555555555555FF5555
        555555555C05555555555555577FF55555555555CCC05555555555557777F555
        55555555CCC05555555555557777FF555555555CCCCC05555555555777777F55
        555555CCCCCC05555555557777777FF5555554CC05CCC05555555777757777F5
        55554C05555CC05555557775555777FF55555555555CCC05555555555557777F
        555555500055CC05555555555555777FF555555060555CC05555555555555777
        FF5555505055554C05555555555555777FF5550556055554C055555555555557
        77FF506F556055555CC055555555555557770000000005555555555555555555
        5555}
      Visible = False
    end
    object CbExo: THValComboBox
      Left = 69
      Top = 8
      Width = 238
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = CbExoChange
      TagDispatch = 0
    end
    object CbVal: TCheckBox
      Left = 336
      Top = 12
      Width = 65
      Height = 17
      Alignment = taLeftJustify
      AllowGrayed = True
      Caption = '&Valid'#233
      State = cbGrayed
      TabOrder = 1
      OnClick = CbValClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 57
    Width = 517
    Height = 24
    Align = alTop
    TabOrder = 1
    object PJal: TPanel
      Left = 1
      Top = 1
      Width = 304
      Height = 22
      Align = alLeft
      Anchors = [akLeft, akTop, akRight, akBottom]
      Caption = 'Journaux'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object PPer: TPanel
      Left = 308
      Top = 1
      Width = 208
      Height = 22
      Align = alRight
      Alignment = taLeftJustify
      Caption = ' P'#233'riode'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
  end
  object CBDebDat: TComboBox
    Left = 436
    Top = 156
    Width = 30
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 2
    Visible = False
  end
  object CBFinDate: TComboBox
    Left = 476
    Top = 156
    Width = 30
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 3
    Visible = False
  end
  object PerJal: TComboBox
    Left = 168
    Top = 30
    Width = 81
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 4
    Visible = False
  end
  object PerJal1: TComboBox
    Left = 264
    Top = 30
    Width = 105
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 5
    Visible = False
  end
  object HPanel1: THPanel
    Left = 0
    Top = 81
    Width = 308
    Height = 229
    Align = alClient
    Caption = 'HPanel1'
    FullRepaint = False
    TabOrder = 7
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Fliste: THGrid
      Left = 1
      Top = 1
      Width = 306
      Height = 227
      Align = alClient
      ColCount = 4
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowSelect]
      TabOrder = 0
      OnDblClick = FlisteDblClick
      OnKeyDown = FlisteKeyDown
      OnMouseDown = FlisteMouseDown
      SortedCol = -1
      Titres.Strings = (
        'Code'
        'Libell'#233
        'Valid'#233)
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
        40
        202
        39
        0)
    end
  end
  object HPanel2: THPanel
    Left = 308
    Top = 81
    Width = 209
    Height = 229
    Align = alRight
    Caption = 'HPanel2'
    FullRepaint = False
    TabOrder = 8
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Fliste1: THGrid
      Left = 1
      Top = 1
      Width = 207
      Height = 227
      Align = alClient
      ColCount = 2
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowSelect]
      TabOrder = 0
      OnClick = Fliste1Click
      SortedCol = -1
      Titres.Strings = (
        'P'#233'riode'
        'Validation')
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
        98
        89)
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Oui'
      'Non'
      'P'#233'riode :'
      'Validation par journal'
      'D'#233'validation par journal'
      'ligne s'#233'lectionn'#233'e'
      'lignes s'#233'lectionn'#233'es'
      'Totale'
      'Partielle'
      'Aucune'
      
        '10;?caption?;Aucun traitement n'#39'est '#224' effectuer pour la p'#233'riode ' +
        'que vous avez choisie.;W;O;O;O;'
      
        '11;?caption?;Ce traitement va valider toutes les p'#233'riodes ant'#233'ri' +
        'eures '#224' celle choisie. D'#233'sirez-vous continuer?;Q;YN;N;N;'
      
        '12;?caption?;Ce traitement va d'#233'valider la p'#233'riode choisie. D'#233'si' +
        'rez-vous continuer?;Q;YN;N;N;'
      '13;?caption?;Vous n'#39'avez s'#233'lectionn'#233' aucun journal.;W;O;O;O;'
      
        '14;?caption?;La p'#233'riode s'#233'lectionn'#233'e a '#233't'#233' clotur'#233'e. Elle ne peu' +
        't '#234'tre d'#233'valid'#233'e.;W;O;O;O;'
      ' ')
    Left = 168
    Top = 188
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 120
    Top = 170
  end
end

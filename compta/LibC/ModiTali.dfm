object FModiTali: TFModiTali
  Left = 328
  Top = 183
  Width = 620
  Height = 424
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Modification en s'#233'rie sur tables libres'
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
    Width = 612
    Height = 121
    ActivePage = PStandards
    Align = alTop
    TabOrder = 0
    object PStandards: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 93
        Align = alClient
      end
      object TCpteDe: TLabel
        Left = 7
        Top = 12
        Width = 112
        Height = 13
        Caption = '&Sections analytiques de'
        FocusControl = CpteDe
      end
      object TCpteA: TLabel
        Left = 264
        Top = 12
        Width = 6
        Height = 13
        Caption = #224
      end
      object TAxe: TLabel
        Left = 431
        Top = 12
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Axe
      end
      object TTL: TLabel
        Left = 7
        Top = 40
        Width = 97
        Height = 13
        Caption = '&Table libre '#224' modifier'
        FocusControl = TL
      end
      object TAncVal: TLabel
        Left = 7
        Top = 68
        Width = 77
        Height = 13
        Caption = 'Ancienne &valeur'
        FocusControl = AncVal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object CpteDeJ: TEdit
        Left = 130
        Top = 8
        Width = 126
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 5
        OnChange = CpteDeChange
      end
      object CpteA: THCpteEdit
        Left = 278
        Top = 8
        Width = 126
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Axe: THValComboBox
        Left = 460
        Top = 8
        Width = 136
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = AxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object TL: THValComboBox
        Left = 122
        Top = 36
        Width = 231
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = TLChange
        TagDispatch = 0
      end
      object CpteDe: THCpteEdit
        Left = 130
        Top = 8
        Width = 126
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        OnChange = CpteDeChange
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object AncVal: THCpteEdit
        Left = 130
        Top = 64
        Width = 126
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 4
        ZoomTable = tzBudgen
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Panel2: TPanel
        Left = 369
        Top = 36
        Width = 227
        Height = 53
        BevelInner = bvLowered
        TabOrder = 6
        object TNewVal: TLabel
          Left = 3
          Top = 20
          Width = 74
          Height = 13
          Caption = '&Nouvelle valeur'
          FocusControl = NewVal
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object NewVal: THCpteEdit
          Left = 85
          Top = 16
          Width = 136
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          ZoomTable = tzGeneral
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
      end
    end
    object PCompGen: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 93
        Align = alClient
      end
      object TNatGen: THLabel
        Left = 32
        Top = 21
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = NatGen
      end
      object TVentil: THLabel
        Left = 337
        Top = 21
        Width = 46
        Height = 13
        Caption = '&Ventilable'
      end
      object NatGen: THValComboBox
        Left = 84
        Top = 17
        Width = 175
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATGENE'
      end
      object Ventil: THValComboBox
        Left = 397
        Top = 17
        Width = 175
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTAXE'
      end
      object G_COLLECTIF: TCheckBox
        Left = 32
        Top = 58
        Width = 61
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Collectif'
        State = cbGrayed
        TabOrder = 2
      end
      object G_FERME: TCheckBox
        Left = 206
        Top = 58
        Width = 53
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Ferm'#233
        State = cbGrayed
        TabOrder = 3
      end
      object G_LETTRABLE: TCheckBox
        Left = 337
        Top = 58
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Lettrable'
        State = cbGrayed
        TabOrder = 4
      end
      object G_POINTABLE: TCheckBox
        Left = 503
        Top = 58
        Width = 69
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Pointable'
        State = cbGrayed
        TabOrder = 5
      end
    end
    object PCompTiers: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel3: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 93
        Align = alClient
      end
      object HLabel1: THLabel
        Left = 280
        Top = 21
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = T_NATUREAUXI
      end
      object TT_COLLECTIF: THLabel
        Left = 280
        Top = 59
        Width = 37
        Height = 13
        Caption = '&Collectif'
        FocusControl = T_COLLECTIF
      end
      object TT_Secteur: THLabel
        Left = 7
        Top = 21
        Width = 37
        Height = 13
        Caption = '&Secteur'
        FocusControl = T_SECTEUR
      end
      object TT_REGIMETVA: THLabel
        Left = 7
        Top = 59
        Width = 60
        Height = 13
        Caption = '&R'#233'gime TVA'
        FocusControl = T_REGIMETVA
      end
      object T_NATUREAUXI: THValComboBox
        Left = 330
        Top = 17
        Width = 176
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATTIERSCPTA'
      end
      object T_COLLECTIF: THCpteEdit
        Left = 330
        Top = 55
        Width = 176
        Height = 21
        MaxLength = 17
        TabOrder = 4
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_SECTEUR: THValComboBox
        Left = 72
        Top = 17
        Width = 186
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTSECTEUR'
      end
      object T_REGIMETVA: THValComboBox
        Left = 72
        Top = 55
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTREGIMETVA'
      end
      object T_LETTRABLE: TCheckBox
        Left = 515
        Top = 21
        Width = 81
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Lettrable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbGrayed
        TabOrder = 2
      end
      object T_MULTIDEVISE: TCheckBox
        Left = 515
        Top = 59
        Width = 81
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Multi devise'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbGrayed
        TabOrder = 5
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 162
    Width = 612
    Height = 191
    Align = alClient
    ColCount = 4
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    TabOrder = 1
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Libell'#233
      'Code table libre')
    Couleur = True
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      120
      285
      184
      0)
  end
  object Dock971: TDock97
    Left = 0
    Top = 353
    Width = 612
    Height = 44
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 40
      ClientWidth = 612
      Caption = 'ToolWindow971'
      ClientAreaHeight = 40
      ClientAreaWidth = 612
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object Tex1: TLabel
        Left = 269
        Top = 11
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
      object Nb1: TLabel
        Left = 240
        Top = 11
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
      object BTag: TToolbarButton97
        Left = 484
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        Flat = False
        Visible = False
        OnClick = BTagClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object Bdetag: TToolbarButton97
        Left = 484
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Tout d'#233's'#233'lectionner'
        Flat = False
        NumGlyphs = 2
        OnClick = BdetagClick
        GlobalIndexImage = 'Z0078_S16G2'
      end
      object BValider: TToolbarButton97
        Left = 517
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Valider'
        Flat = False
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
          A400000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          0303030303030303030303030303030303030303030303030303030303030303
          03030303030303030303030303030303030303030303FF030303030303030303
          03030303030303040403030303030303030303030303030303F8F8FF03030303
          03030303030303030303040202040303030303030303030303030303F80303F8
          FF030303030303030303030303040202020204030303030303030303030303F8
          03030303F8FF0303030303030303030304020202020202040303030303030303
          0303F8030303030303F8FF030303030303030304020202FA0202020204030303
          0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
          040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
          03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
          FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
          0303030303030303030303FA0202020403030303030303030303030303F8FF03
          03F8FF03030303030303030303030303FA020202040303030303030303030303
          0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
          03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
          030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
          0202040303030303030303030303030303F8FF03F8FF03030303030303030303
          03030303FA0202030303030303030303030303030303F8FFF803030303030303
          030303030303030303FA0303030303030303030303030303030303F803030303
          0303030303030303030303030303030303030303030303030303030303030303
          0303}
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        OnClick = BValiderClick
      end
      object BFerme: TToolbarButton97
        Left = 549
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Flat = False
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082840082840082840082840082840082848482848482840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284FFFFFF008284008284008284008284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          840000848482840082840082840082840082840082840000FF84828400828400
          8284008284008284008284008284008284008284848284848284FFFFFF008284
          008284008284008284008284008284FFFFFF0082840082840082840082840082
          840082840082840000FF00008400008400008484828400828400828400828400
          00FF000084000084848284008284008284008284008284008284008284848284
          FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
          FF0082840082840082840082840082840082840000FF00008400008400008400
          00848482840082840000FF000084000084000084000084848284008284008284
          008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
          FF848284008284008284848284FFFFFF00828400828400828400828400828400
          82840000FF000084000084000084000084848284000084000084000084000084
          000084848284008284008284008284008284008284848284FFFFFF0082840082
          84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
          82840082840082840082840082840082840000FF000084000084000084000084
          0000840000840000840000848482840082840082840082840082840082840082
          84008284848284FFFFFF00828400828400828484828400828400828400828400
          8284FFFFFF848284008284008284008284008284008284008284008284008284
          0000FF0000840000840000840000840000840000848482840082840082840082
          84008284008284008284008284008284008284848284FFFFFF00828400828400
          8284008284008284008284FFFFFF848284008284008284008284008284008284
          0082840082840082840082840082840000840000840000840000840000848482
          8400828400828400828400828400828400828400828400828400828400828400
          8284848284FFFFFF008284008284008284008284008284848284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          8400008400008400008484828400828400828400828400828400828400828400
          8284008284008284008284008284008284848284FFFFFF008284008284008284
          8482840082840082840082840082840082840082840082840082840082840082
          840082840000FF00008400008400008400008400008484828400828400828400
          8284008284008284008284008284008284008284008284008284008284848284
          008284008284008284008284848284FFFFFF0082840082840082840082840082
          840082840082840082840082840000FF00008400008400008484828400008400
          0084000084848284008284008284008284008284008284008284008284008284
          008284008284848284008284008284008284008284008284848284FFFFFF0082
          840082840082840082840082840082840082840082840000FF00008400008400
          00848482840082840000FF000084000084000084848284008284008284008284
          008284008284008284008284008284848284008284008284008284848284FFFF
          FF008284008284848284FFFFFF00828400828400828400828400828400828400
          82840000FF0000840000848482840082840082840082840000FF000084000084
          000084848284008284008284008284008284008284008284848284FFFFFF0082
          84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
          82840082840082840082840082840082840000FF000084008284008284008284
          0082840082840000FF0000840000840000840082840082840082840082840082
          84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
          8284008284848284FFFFFF008284008284008284008284008284008284008284
          0082840082840082840082840082840082840082840000FF0000840000FF0082
          8400828400828400828400828400828400828484828484828400828400828400
          8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284848284848284848284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          008284008284008284008284008284008284}
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        OnClick = BFermeClick
      end
      object BAide: TToolbarButton97
        Left = 581
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Aide'
        Flat = False
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 121
    Width = 612
    Height = 41
    AllowDrag = False
    object ToolWindow972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 37
      ClientWidth = 612
      Caption = 'ToolWindow972'
      ClientAreaHeight = 37
      ClientAreaWidth = 612
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BChercher: TToolbarButton97
        Left = 568
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Flat = False
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
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
      object FFiltres: THValComboBox
        Left = 68
        Top = 7
        Width = 492
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '&Sections analytiques de'
      '&Comptes g'#233'n'#233'raux de'
      '&Comptes auxiliaires de'
      '&Comptes budg'#233'taires'
      '&Sections budg'#233'taires de'
      
        '5;Modification en s'#233'rie des tables libres;Le code renseign'#233' n'#39'ex' +
        'iste pas;W;O;O;O;'
      'ligne s'#233'lectionn'#233'e'
      'lignes s'#233'lectionn'#233'es'
      
        '8;Modification en s'#233'rie des tables libres;Aucun compte ne corres' +
        'pond '#224' votre crit'#232're, aucun traitement '#224' effectuer;W;O;O;O;'
      
        '9;Modification en s'#233'rie des tables libres;Aucun compte n'#39'est s'#233'l' +
        'ectionn'#233', aucun traitement a effectuer;W;O;O;O;'
      
        '10;Modification en s'#233'rie des tables libres;Confirmez-vous le tra' +
        'itement de mise '#224' jour des tables libres s'#233'lectionn'#233'es?;Q;YN;N;N' +
        ';'
      
        '11;Modification en s'#233'rie des tables libres;La nouvelle valeur es' +
        't '#233'gale '#224' l'#39'ancienne valeur. Aucun traitement '#224' effectuer !;W;O;' +
        'O;O;'
      
        '12;Modification en s'#233'rie des tables libres;Des enregistrements n' +
        #39'ont pas '#233't'#233' mis '#224' jour, il n'#39'ont pas pu '#234'tre localis'#233's.;W;O;O;O' +
        ';'
      '&Immobilisation de'
      '14;')
    Left = 372
    Top = 220
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 296
    Top = 194
  end
  object POPF: TPopupMenu
    Left = 191
    Top = 236
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

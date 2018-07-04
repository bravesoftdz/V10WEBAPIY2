object FCreSecBu: TFCreSecBu
  Left = 242
  Top = 169
  Width = 587
  Height = 400
  HelpContext = 7577430
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 
    'G'#233'n'#233'ration des sections budg'#233'taires '#224' partir d'#39'un plan de sectio' +
    'ns'
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
    Width = 579
    Height = 97
    ActivePage = Param
    Align = alTop
    TabOrder = 0
    object Param: TTabSheet
      Caption = 'Param'#233'trage'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 571
        Height = 69
        Align = alClient
      end
      object THLabel
        Left = 31
        Top = 22
        Width = 3
        Height = 13
      end
      object TSigne: THLabel
        Left = 230
        Top = 4
        Width = 27
        Height = 13
        Caption = 'S&igne'
        Color = clYellow
        FocusControl = Signe
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object TSens: THLabel
        Left = 283
        Top = 18
        Width = 24
        Height = 13
        Caption = '&Sens'
        FocusControl = Sens
      end
      object TAxe: THLabel
        Left = 16
        Top = 18
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Axe
      end
      object Nb1: TLabel
        Left = 63
        Top = 45
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
        Left = 88
        Top = 45
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
        Left = 531
        Top = 11
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object Sens: THValComboBox
        Left = 319
        Top = 14
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTSENS'
      end
      object Signe: THValComboBox
        Left = 271
        Top = 0
        Width = 27
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
        TabOrder = 1
        Visible = False
        TagDispatch = 0
        DataType = 'TTRUBSIGNE'
      end
      object Axe: THValComboBox
        Left = 52
        Top = 14
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = AxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 338
    Width = 579
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 575
      Caption = 'Actions'
      ClientAreaHeight = 31
      ClientAreaWidth = 575
      DockPos = 0
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      DesignSize = (
        575
        31)
      object BTag: TToolbarButton97
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'S'#233'lectionner'
        Flat = False
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000130B0000130B00001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          333333333333333333FF33333333333330003333333333333888333333333333
          300033FFFFFF3333388839999993333333333888888F3333333F399999933333
          3300388888833333338833333333333333003333333333333388333333333333
          3333333333333333333F3333333E3333330033333F33333333883333333EC333
          330033338F3333333388EEEEEEEECCE3333333F88FFFFFFF3FF3ECCCCCCCCCCC
          399338888888888F88F33CCCCCCCCCCC399338888888888388333333333ECC33
          333333388F33333333FF33333333C33330003333833333333888333333333333
          3000333333333333388833333333333333333333333333333333}
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BTagClick
      end
      object BCodeAbr: TToolbarButton97
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Codes abr'#233'g'#233's des plans de sous sections'
        DisplayMode = dmGlyphOnly
        Caption = 'Codes'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BCodeAbrClick
        GlobalIndexImage = 'Z0210_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 480
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 512
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
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
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
      end
      object BAide: TToolbarButton97
        Left = 544
        Top = 2
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
        Top = 2
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
  object FListe: THGrid
    Left = 0
    Top = 97
    Width = 579
    Height = 241
    Align = alClient
    ColCount = 4
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    TabOrder = 2
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Libell'#233
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
      175
      260
      120
      0)
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Ligne s'#233'lectionn'#233'e'
      'Lignes s'#233'lectionn'#233'es'
      
        '2;G'#233'n'#233'ration des sections budg'#233'taires;Les sections non mouvement' +
        #233'es vont '#234'tre d'#233'truites et remplac'#233'es. Confirmez-vous le traitem' +
        'ent ?;Q;YN;N;N;'
      'Destruction'
      'Cr'#233'ation'
      
        '5;G'#233'n'#233'ration des sections budg'#233'taires;Vous devez param'#233'trer la c' +
        'opie des sections analytiques vers les sections budg'#233'taires.;W;O' +
        ';O;O;'
      
        '6;G'#233'n'#233'ration des sections budg'#233'taires;Des sections budg'#233'taires n' +
        #39'ont pas '#233't'#233' cr'#233#233'es. D'#233'sirez-vous les modifier pour les cr'#233#233'r ?;' +
        'Q;YNC;Y;N;')
    Left = 412
    Top = 164
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 180
  end
end

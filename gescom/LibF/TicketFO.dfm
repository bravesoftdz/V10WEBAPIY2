object FTicketFO: TFTicketFO
  Left = 119
  Top = 263
  Width = 726
  Height = 518
  HelpContext = 301100240
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSizeToolWin
  Caption = 'Saisie de pièce'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object PMain: TPanel
    Left = 0
    Top = 86
    Width = 718
    Height = 367
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 12
    object ToolbarButton972: TToolbarButton97
      Left = 396
      Top = 0
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      DisplayMode = dmGlyphOnly
      Caption = 'Imprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Glyph.Data = {
        4E010000424D4E01000000000000760000002800000012000000120000000100
        040000000000D800000000000000000000001000000010000000000000000000
        BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
        DDDDDD000000DDD00000000000DDDD000000DD0777777777070DDD000000D000
        000000000070DD000000D0777777FFF77000DD000000D077777799977070DD00
        0000D0000000000000770D000000D0777777777707070D000000DD0000000000
        70700D000000DDD0FFFFFFFF07070D000000DDDD0FCCCCCF0000DD000000DDDD
        0FFFFFFFF0DDDD000000DDDDD0FCCCCCF0DDDD000000DDDDD0FFFFFFFF0DDD00
        0000DDDDDD000000000DDD000000DDDDDDDDDDDDDDDDDD000000DDDDDDDDDDDD
        DDDDDD000000DDDDDDDDDDDDDDDDDD000000}
      Layout = blGlyphTop
      Margin = 2
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Visible = False
      OnClick = BImprimerClick
    end
    object PGrille: TPanel
      Left = 1
      Top = 1
      Width = 716
      Height = 229
      Align = alTop
      TabOrder = 1
      object PClavier2: TPanel
        Left = 356
        Top = 1
        Width = 359
        Height = 227
        TabOrder = 1
        Visible = False
      end
      object PnlCorps: THPanel
        Left = 1
        Top = 1
        Width = 714
        Height = 227
        Align = alClient
        BevelOuter = bvNone
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object PPied: THPanel
          Left = 0
          Top = 143
          Width = 714
          Height = 84
          Align = alBottom
          Color = clSilver
          Ctl3D = True
          FullRepaint = False
          ParentShowHint = False
          ShowHint = False
          TabOrder = 0
          BackGroundEffect = bdHorzIn
          ColorShadow = clWindowText
          ColorStart = clSilver
          TextEffect = tenone
          object ImageArticle: TImage
            Left = 4
            Top = 4
            Width = 75
            Height = 75
            Stretch = True
            Transparent = True
          end
          object LblPrixU: THLabel
            Left = 5
            Top = 18
            Width = 68
            Height = 16
            Caption = 'Prix unitaire'
            Color = clInfoBk
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
            Transparent = True
          end
          object LblA_LIBELLE: THLabel
            Left = 5
            Top = 33
            Width = 34
            Height = 16
            Caption = 'Stock'
            Color = clInfoBk
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
            Transparent = True
          end
          object T_PrixU: TFlashingLabel
            Tag = 2
            Left = 123
            Top = 18
            Width = 125
            Height = 19
            Hint = 'P_'
            Alignment = taRightJustify
            AutoSize = False
            BiDiMode = bdLeftToRight
            Caption = '0,00'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentBiDiMode = False
            ParentFont = False
            Transparent = True
          end
          object A_QTESTOCK: TFlashingLabel
            Left = 123
            Top = 33
            Width = 125
            Height = 19
            Alignment = taRightJustify
            AutoSize = False
            Caption = '0,00'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            Transparent = True
          end
          object PI_NETAPAYERDEV: TLCDLabel
            Left = 386
            Top = 6
            Width = 322
            Height = 46
            Caption = '        0.00'
            PixelSize = pix4x4
            PixelSpacing = 1
            CharSpacing = 2
            LineSpacing = 2
            BorderSpace = 5
            TextLines = 1
            NoOfChars = 12
            BackGround = clBlack
            PixelOn = clLime
            PixelOff = clBlack
            UpSideDown = False
          end
          object TPI_NETAPAYERDEV: THLabel
            Left = 319
            Top = 14
            Width = 60
            Height = 29
            Caption = 'Total'
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -24
            Font.Name = 'Microsoft Sans Serif'
            Font.Style = [fsBold]
            ParentFont = False
            Transparent = True
            Ombre = True
            OmbreDecalX = 1
            OmbreDecalY = 1
            OmbreColor = clWhite
          end
          object LIBELLETIERS: THLabel
            Left = 4
            Top = 2
            Width = 365
            Height = 18
            AutoSize = False
            Caption = 'LIBELLETIERS'
            Color = clInfoBk
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            ParentColor = False
            ParentFont = False
            Transparent = True
          end
          object LIBELLEARTICLE: THLabel
            Left = 4
            Top = 49
            Width = 289
            Height = 16
            AutoSize = False
            Caption = 'LIBELLEARTICLE'
            Color = clInfoBk
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
            Transparent = True
          end
          object LGP_TOTALQTEFACT: TLabel
            Left = 462
            Top = 57
            Width = 111
            Height = 13
            Caption = 'LGP_TOTALQTEFACT'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            Transparent = True
          end
          object TLGP_TOTALQTEFACT: TLabel
            Left = 386
            Top = 57
            Width = 75
            Height = 13
            Caption = 'Quantité totale :'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            Transparent = True
          end
          object LDETAXE: TLabel
            Left = 667
            Top = 57
            Width = 41
            Height = 13
            Caption = 'Détaxe'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            ParentFont = False
            Transparent = True
            Visible = False
          end
          object SOLDETIERS: TFlashingLabel
            Tag = 2
            Left = 123
            Top = 63
            Width = 125
            Height = 19
            Hint = 'P_'
            Alignment = taRightJustify
            AutoSize = False
            BiDiMode = bdLeftToRight
            Caption = '0,00'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentBiDiMode = False
            ParentFont = False
            Transparent = True
          end
          object TSOLDETIERS: THLabel
            Left = 5
            Top = 64
            Width = 104
            Height = 20
            Caption = 'Solde du client'
            Color = clInfoBk
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
            Transparent = True
          end
          object LALIVRER: TLabel
            Left = 319
            Top = 57
            Width = 41
            Height = 13
            Caption = 'A livrer'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            ParentFont = False
            Transparent = True
            Visible = False
          end
          object GP_TOTALHTDEV: THNumEdit
            Tag = 1
            Left = 475
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Decimals = 2
            Digits = 12
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentFont = False
            TabOrder = 0
            UseRounding = True
            Validate = False
            Visible = False
            OnChange = AfficheZonePied
          end
          object GP_TOTALREMISEDEV: THNumEdit
            Tag = 1
            Left = 512
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Decimals = 2
            Digits = 12
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentFont = False
            TabOrder = 1
            UseRounding = True
            Validate = False
            Visible = False
            OnChange = AfficheZonePied
          end
          object HTOTALTAXESDEV: THNumEdit
            Tag = 1
            Left = 549
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Decimals = 2
            Digits = 12
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentFont = False
            TabOrder = 2
            UseRounding = True
            Validate = False
            Visible = False
            OnChange = AfficheZonePied
          end
          object GP_TOTALESCDEV: THNumEdit
            Tag = 1
            Left = 585
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Decimals = 2
            Digits = 12
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentFont = False
            TabOrder = 3
            UseRounding = True
            Validate = False
            Visible = False
            OnChange = AfficheZonePied
          end
          object GP_TOTALTTCDEV: THNumEdit
            Tag = 1
            Left = 622
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Decimals = 2
            Digits = 12
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentFont = False
            TabOrder = 4
            UseRounding = True
            Validate = False
            Visible = False
            OnChange = AfficheZonePied
          end
          object GP_ESCOMPTE: THNumEdit
            Left = 662
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Ctl3D = True
            Decimals = 2
            Digits = 15
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 5
            UseRounding = True
            Validate = False
            Visible = False
            OnExit = GP_ESCOMPTEExit
          end
          object GP_REMISEPIED: THNumEdit
            Left = 397
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Ctl3D = True
            Decimals = 2
            Digits = 15
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 6
            UseRounding = True
            Validate = False
            Visible = False
            OnExit = GP_REMISEPIEDExit
          end
          object GP_MODEREGLE: THValComboBox
            Left = 397
            Top = 28
            Width = 40
            Height = 21
            Style = csSimple
            Color = clYellow
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 13
            ParentFont = False
            TabOrder = 7
            Text = 'GP_MODEREGLE'
            Visible = False
            TagDispatch = 0
            DataType = 'TTMODEREGLE'
          end
          object GP_TOTALQTEFACT: THNumEdit
            Tag = 5
            Left = 436
            Top = 8
            Width = 40
            Height = 21
            Color = clYellow
            Decimals = 2
            Digits = 12
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clRed
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Masks.PositiveMask = '#,##0.00'
            Debit = False
            ParentFont = False
            TabOrder = 8
            UseRounding = True
            Validate = False
            Visible = False
            OnChange = AfficheZonePied
          end
        end
        object GS: THGrid
          Tag = 1
          Left = 0
          Top = 0
          Width = 714
          Height = 143
          Align = alClient
          ColCount = 8
          Ctl3D = False
          DefaultColWidth = 50
          DefaultRowHeight = 23
          FixedColor = clWindow
          RowCount = 6
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goTabs]
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 1
          OnDblClick = GSDblClick
          OnEnter = GSEnter
          OnMouseDown = GSMouseDown
          SortedCol = -1
          Couleur = False
          MultiSelect = False
          TitleBold = True
          TitleCenter = True
          OnRowEnter = GSRowEnter
          OnRowExit = GSRowExit
          OnCellEnter = GSCellEnter
          OnCellExit = GSCellExit
          ColCombo = 0
          SortEnabled = False
          SortRowExclude = 0
          TwoColors = False
          AlternateColor = 13224395
          OnElipsisClick = GSElipsisClick
          ColWidths = (
            19
            16
            108
            209
            78
            68
            55
            82)
          RowHeights = (
            23
            23
            23
            23
            23
            23)
        end
      end
    end
    object PImage: THPanel
      Left = 244
      Top = 248
      Width = 413
      Height = 117
      FullRepaint = False
      TabOrder = 0
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 0
      TextEffect = tenone
    end
  end
  object DockBottom: TDock97
    Left = 0
    Top = 453
    Width = 718
    Height = 31
    BackgroundTransparent = True
    Position = dpBottom
    object Outils97: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Actions'
      CloseButton = False
      DefaultDock = DockBottom
      DockPos = 0
      TabOrder = 0
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 0
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Menu zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Caption = 'Zoom'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          AA040000424DAA04000000000000360000002800000013000000130000000100
          1800000000007404000000000000000000000000000000000000FF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FFFFFFFFFFFFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFBF0000BF
          0000FFFFFFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFBF000000FFFFFFFF00BF00
          00FFFFFFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFFFFFFBF000000FFFFFFFF80FF8000BF0000FF00FF
          FF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFBF000000FFFFFFFF00BF0000BF0000FF00FFFF00FFFF00FF00
          0000FF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFF000000000000000000000000
          BF000000FFFFFFFF00FF8000BF0000FF00FFFF00FFFF00FFFF00FF000000FF00
          FFFF00FFFFFFFFFFFFFF000000000000FF00FFFF00FFFF00FFFF00FFFF00FFBF
          0000BF0000FF0040FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
          FFFFFF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0000
          00FF0040FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFFFFFFFF00
          0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF
          00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00FF00FFFF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00
          FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00FF00FFFFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FF
          FF00FF000000FF00FFFF00FF000000FF00FFFF00FF00FFFF00FFFFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FF00
          0000FF00FFFF00FFFF00FF000000FF00FFFF00FF00FFFF00FFFF00FFFFFF00FF
          FF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00
          FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
          FF00FFFF00FF000000000000FF00FFFF00FFFF00FFFF00FF000000000000FF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FF000000000000000000000000FF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FF000000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomMouseEnter
        IsControl = True
      end
      object BEche: TToolbarButton97
        Tag = 1
        Left = 206
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Echéances'
        DisplayMode = dmGlyphOnly
        Caption = 'Echéances'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          06020000424D0602000000000000760000002800000028000000140000000100
          0400000000009001000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          33333333333333333333333333333333333333333333333333FFFFFFFFFFFF33
          3333300000000000033333333777777777777F33333330F88888888803333333
          37F3333333337F33333330F8888888880333333337F3333333337FFFF33330F8
          888888880000333337F3333333337777F33330F8888888880880333337F33333
          33337F37F33330F8888888880880333337F3333333337F37FFF330F888888888
          0880000337F3333333337F37777F30FFFFFFFFFF0880880337FFFFFFFFFF7F37
          F37F304444444444088088033777777777777F37F37F30000000000008808803
          3777777777777337F37F33330FFFFFFFFFF0880333337FFFFFFFFFF7F37F3333
          04444444444088033333777777777777F37F3333000000000000880333337777
          77777777337F33333330FFFFFFFFFF0333333337FFFFFFFFFF7F333333304444
          444444033333333777777777777F333333300000000000033333333777777777
          7773333333333333333333333333333333333333333333333333333333333333
          33333333333333333333}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BEcheClick
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 40
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions complémentaires'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPY
        Caption = 'Complément'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          06020000424D0602000000000000760000002800000028000000140000000100
          0400000000009001000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333FFFFFFFFFF333333333000000000033333333337777777777F3333
          333330FEFEFEFE033333333337F3FFFFFF7F3333333330E444444F0333333333
          37F77777737F333333333099999999033333333337F33333337F33333333309F
          FFFFF9033333333337F33333337F333333333099999999033333333337F3FFFF
          FF7F3333333330E444444F033333333337F77777737F3333333330FEFEFEFE03
          3333333337F3FFF3FF7F3333333330E444E000033333333337F7773777733333
          333330FEFEF0F0333333333337F33337F7333333333330EFEFE0033333333333
          37FFFFF7733333FFFF333000000033333000033337777777333337777F333333
          3333333330EA0333333333333333F7F37FFF33333333333000AE000333333333
          33377733777F333333333330EAEAEA03333333333337FFF33F7F333333333330
          00AE000333333333333777F3777333333333333330EA033333333333333337FF
          7F33333333333333300003333333333333333777733333333333333333333333
          33333333333333333333}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 432
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la pièce'
        DisplayMode = dmGlyphOnly
        Caption = 'Rechercher'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC00000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000077777777777777777000000070000077777000007000000070B0
          00777770F0007000000070F000777770B0007000000070000000700000007000
          0000700B000000B0000070000000700F000700F0000070000000700B000700B0
          0000700000007700000000000007700000007770B00070B00077700000007770
          0000700000777000000077770007770007777000000077770B07770B07777000
          0000777700077700077770000000777777777777777770000000777777777777
          777770000000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BChercherClick
        IsControl = True
      end
      object BSousTotal: TToolbarButton97
        Tag = 1
        Left = 320
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Insérer un sous-total'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'Sous-total'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          0400000000000001000000000000000000001000000010000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          333333333333333333333333333333333333333FFFFFFFFFFF33330000000000
          03333377777777777F33333003333330033333377FF333377F33333300333333
          0333333377FF33337F3333333003333303333333377FF3337333333333003333
          333333333377FF3333333333333003333333333333377FF33333333333330033
          3333333333337733333333333330033333333333333773333333333333003333
          33333333337733333F3333333003333303333333377333337F33333300333333
          03333333773333337F33333003333330033333377FFFFFF77F33330000000000
          0333337777777777733333333333333333333333333333333333}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BSousTotalClick
        IsControl = True
      end
      object BActionsLignes: TToolbarButton97
        Tag = 1
        Left = 280
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions lignes'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopL
        Caption = 'Lignes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          0400000000000001000000000000000000001000000010000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          33333FFFFFFFFFFFFFFF000000000000000077777777777777770FFFFFFFFFFF
          FFF07F3FF3FF3FFF3FF70F00F00F000F00F07F773773777377370FFFFFFFFFFF
          FFF07F3FF3FF33FFFFF70F00F00FF00000F07F773773377777F70FEEEEEFF0F9
          FCF07F33333337F7F7F70FFFFFFFF0F9FCF07F3FFFF337F737F70F0000FFF0FF
          FCF07F7777F337F337370F0000FFF0FFFFF07F777733373333370FFFFFFFFFFF
          FFF07FFFFFFFFFFFFFF70CCCCCCCCCCCCCC07777777777777777088CCCCCCCCC
          C880733777777777733700000000000000007777777777777777333333333333
          3333333333333333333333333333333333333333333333333333}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
      object ToolbarSep971: TToolbarSep97
        Left = 200
        Top = 0
      end
      object Sep2: TToolbarSep97
        Left = 274
        Top = 0
      end
      object BOffice: TToolbarButton97
        Tag = 1
        Left = 516
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Edition bureautique'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'Bureautique'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000014000000120000000100
          040000000000D800000000000000000000001000000010000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7777777700007777777777777777777700007777777777777777777700007777
          7777777777777777000077777777777777777777000077700077000777777777
          00007780EE080EE077777777000077806EE006EE07777777000077780EEE06EE
          E07777770000777806E0E0E60E0777770000777780E00EEE00E0777700007770
          00EE00EE600E077700007700EEEE600EE0EEE077000077000000070000000077
          0000770000007770070007770000777777777777777777770000777777777777
          777777770000777777777777777777770000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BOfficeClick
        IsControl = True
      end
      object BVentil: TToolbarButton97
        Tag = 1
        Left = 234
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Ventilations analytiques'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopV
        Caption = 'Analytiques'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          66010000424D6601000000000000760000002800000014000000140000000100
          040000000000F000000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00888888888888
          8888888800008888888888888888888800008888888888888888888800008888
          8884444488888888000088888884FFF488888888000088888804444488888888
          00008888808888888888888800004444488444448844444800004FFF4004BBB4
          004FFF4800004444488444448844444800008888808888888888888800008888
          8804444488444448000088888884FFF4004FFF48000088888884444488444448
          0000888888888888088888880000888888888888804444480000888888888888
          884FFF4800008888888888888844444800008888888888888888888800008888
          88888888888888880000}
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
      object BDescriptif: TToolbarButton97
        Tag = 1
        Left = 488
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Descriptif détaillé de l'#39'article'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = 'Descriptif'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          F6000000424DF600000000000000760000002800000010000000100000000100
          0400000000008000000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7777777777777777707777777777777707777777777007707777777777077007
          777777777707F707777777777707FF7077777777707FFFF07777777707FF0007
          777777007FF07777777770F70F077777777770FF707777777777770FF0777777
          7777777007777777777777777777777777777777777777777777}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BDescriptifClick
        IsControl = True
      end
      object BAcompte: TToolbarButton97
        Tag = 1
        Left = 376
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Saisir un acompte'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'Acompte'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          F6000000424DF600000000000000760000002800000010000000100000000100
          0400000000008000000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00222222222222
          222200000000000022220FFFFFFFFFF022220F00F00000F022220FFFFFFFFFF0
          22220F00F00000F022220FFFFFFFFFF022220FFFFFFF0FF022220F00FFF080F0
          22220F080F08080002440FF08080808880440000080808888844222220808888
          8844222222088888804422222220000002442222222222222222}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        IsControl = True
      end
      object BDelete: TToolbarButton97
        Left = 460
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Supprimer la pièce'
        DisplayMode = dmGlyphOnly
        Caption = 'Supprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          66010000424D6601000000000000760000002800000014000000140000000100
          040000000000F000000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00888888888888
          8888888800008888888888888888889800008898888888888888898800008899
          88777777777798880000889990000000000998880000888990BFFFBFFF998888
          0000888899FCCCCCCF97888800008888999FBFFFB9978888000088888999CCC9
          990788880000888880999FB99F0788880000888880FC9999CF07888800008888
          80FF9999BF0788880000888880FC9999000788880000888880B99F099F078888
          0000888880999F099998888800008888999FBF0F089988880000889999000000
          8889988800008899988888888888898800008888888888888888889800008888
          88888888888888880000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BDeleteClick
      end
      object BPorcs: TToolbarButton97
        Tag = 1
        Left = 404
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Ports et frais'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'Port'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC00000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FFFFFFFFFFFF
          FFFFF0000000FFFFFFFF00000000F0000000FFFFFFFF0FBFBFB0F0000000FFFF
          F0000BFBFBF0F0000000FFFFF0FF00000000F0000000FFFFF0FFFFF0FFFFF000
          0000FFF8F0F888F0FFFFF0000000FFF8B0FFFFF0FFFFF00000008BF8F08F88F0
          FFFFF0000000F8F8B8BFFFF0FFFFF0000000FF8BFB000000FFFFF0000000F88F
          BFFFFFFFFFFFF0000000F8F8F8BFFFFFFFFFF00000008BF8BFFFFFFFFFFFF000
          0000FFF8FFF8FFFFFFFFF0000000FFF8BFFFFFFFFFFFF0000000FFFFFFFFFFFF
          FFFFF0000000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BPorcsClick
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 544
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000012000000120000000100
          040000000000D800000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
          DDDDDD000000DDD00000000000DDDD000000DD0777777777070DDD000000D000
          000000000070DD000000D0777777FFF77000DD000000D077777799977070DD00
          0000D0000000000000770D000000D0777777777707070D000000DD0000000000
          70700D000000DDD0FFFFFFFF07070D000000DDDD0FCCCCCF0000DD000000DDDD
          0FFFFFFFF0DDDD000000DDDDD0FCCCCCF0DDDD000000DDDDD0FFFFFFFF0DDD00
          0000DDDDDD000000000DDD000000DDDDDDDDDDDDDDDDDD000000DDDDDDDDDDDD
          DDDDDD000000DDDDDDDDDDDDDDDDDD000000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BImprimerClick
      end
      object BClient: TToolbarButton97
        Tag = -100
        Left = 80
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions sur les clients'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopT
        Caption = 'Client'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000014000000120000000100
          040000000000D800000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7777777700007777777777777777777700007777770077777777777700007777
          7077007777777777000077777077770077777777000077770770077700777777
          00007777070FF0077700777700007770770FFFF0077700770000777070F000FF
          F00770770000770770F8DD0FFF0770770000770770FF88F88F07077700007700
          000000F8F077077700007770FFFFF00FF070777700007770F0EEF07007707777
          00007770F000F0777707777700007770FFFFF007770777770000777000000070
          007777770000777777777777777777770000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
      object BTicketAttente: TToolbarButton97
        Tag = 1
        Left = 348
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Reprendre un ticket en attente'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'En attente'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000120B0000120B00001000000000000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333000000000
          3333333777777777F3333330F77777703333333733F3F3F73F33330FF0808077
          0333337F37F7F7F37F33330FF0807077033333733737F73F73F330FF77808707
          703337F37F37F37F37F330FF08807807703037F37F37F37F37F700FF08808707
          700377F37337F37F377330FF778078077033373F73F7F3733733330FF0808077
          0333337F37F7F7F37F33330FF08070770333337FF7F7F7FF7F33330000000000
          03333377777777777F33330F888777770333337FFFFFFFFF7F33330000000000
          033333777777777773333333307770333333333337FFF7F33333333330000033
          3333333337777733333333333333333333333333333333333333}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BTicketAttenteClick
        IsControl = True
      end
      object BActionsDiv: TToolbarButton97
        Tag = -100
        Left = 120
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions diverses'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopM
        Caption = 'Divers'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000014000000120000000100
          040000000000D800000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777777000077700000000000000777000077707778FF7FF7FF077700007770
          77788F78F78F07770000777088888778778707770000777077780078F78F0777
          0000777077780E0FF78F077700007770888870E077770777000077700000FF0E
          07FF07770000777077770F70E0FF0777000077077777707F0E0F07770000770F
          7555707FF0E007770000770F75777044440E07770000770F757770000000E077
          00007770FFF70777777700770000777700007777777777770000777777777777
          777777770000777777777777777777770000}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomMouseEnter
        IsControl = True
      end
      object BTaxation: TToolbarButton97
        Tag = 1
        Left = 160
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions sur la taxation'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopX
        Caption = 'Taxes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          18000000000000030000000000000000000000000000000000003F3F3F7F7F3F
          7F7F3F7F7F3F7F7F3F7F7F00000000FFFFFFFFFFFF0000007F007F6F306F7F3F
          7F7F3F7F7F3F7F3F3F3F7F7F3FFFFF7FFFFF7FFFFF7FFFFF3FEFEF3020000060
          0000600000200000DF20DFEF30EFEF6FEFFF7FFFFF7FFF7F3F7F7F7F3FFFFF7F
          FFFF7FFFFF3FEFEF30CFCF30400000DFBFBF800000400000BF20BFDF20DFEF30
          EFEF6FEFFF7FFF7F3F7F7F7F3FFFFF7FFFFF3FEFEF30CFCF30202000BFBFBFCF
          CFCFBF7F7F5F3F3FBF00BFBF20BFDF20DFEF30EFEF6FEF7F3F7F7F7F3FFFFF3F
          EFEF30CFCF307F7F003F3F3FDFDFDFC0C0C0C0C0C0606060BF00BF7F007FBF20
          BFDF20DFEF30EF6F306F7F7F00EFEF30CFCF30BFBF00BFBF006060003F3F3FCF
          CFCF2000002000006000603F3F3F200020BF20BFDF20DF7F007F000000200000
          4000005F3F3F6060602000006000005F3F3F4000006000003F3F3FDFDFDFBFBF
          BF400000200000000000FFFFFF600000800000BF7F7FC0C0C020000040000080
          00008000005F3F3FCFCFCFC0C0C0CFCFCFDFBFBF600000FFFFFFFFFFFF600000
          DFBFBFCFCFCFC0C0C0CFCFCF5F3F3F800000800000400000200000C0C0C0BF7F
          7F800000600000FFFFFF000000200000400000BFBFBFDFDFDF3F3F3F60000040
          00005F3F3F4000002000006060605F3F3F400000200000000000007F0030EF30
          30CF300020003F3F3F006000200000200000CFCFCF5F3F3F00606000BFBF00BF
          BF30CFCF30EFEF007F7F3F7F3F3FFF3F30EF3030CF30007F0000BF00606060C0
          C0C0C0C0C0DFDFDF3F3F3F007F7F30CFCF30EFEF3FFFFF3F7F7F3F7F3F7FFF7F
          3FFF3F30EF3030CF3000BF005F3F3FBF7F7FCFCFCFBFBFBF00202030CFCF30EF
          EF3FFFFF7FFFFF3F7F7F3F7F3F7FFF7F7FFF7F3FFF3F30EF3030CF3040000080
          0000DFBFBF40000030CFCF30EFEF3FFFFF7FFFFF7FFFFF3F7F7F3F7F3F7FFF7F
          7FFF7F7FFF7F3FFF3F30EF3020000060000060000020000030EFEF3FFFFF7FFF
          FF7FFFFF7FFFFF3F7F7F3F3F3F3F7F3F3F7F3F3F7F3F3F7F3F007F00000000FF
          FFFFFFFFFF000000007F7F3F7F7F3F7F7F3F7F7F3F7F7F3F3F3F}
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
    end
    object Valide97: TToolbar97
      Left = 609
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DefaultDock = DockBottom
      DockPos = 609
      TabOrder = 1
      object BValider: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Enregistrer la pièce'
        DisplayMode = dmGlyphOnly
        Caption = 'Enregistrer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080008080008080008080
          0080800080800080800080800080800080808080808080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080FFFFFF008080008080008080008080008080008080
          0080800080800080800080800080800080800080800080800080800000FF0000
          800000808080800080800080800080800080800080800000FF80808000808000
          8080008080008080008080008080008080008080808080808080FFFFFF008080
          008080008080008080008080008080FFFFFF0080800080800080800080800080
          800080800080800000FF00008000008000008080808000808000808000808000
          00FF000080000080808080008080008080008080008080008080008080808080
          FFFFFF008080808080FFFFFF008080008080008080FFFFFF808080808080FFFF
          FF0080800080800080800080800080800080800000FF00008000008000008000
          00808080800080800000FF000080000080000080000080808080008080008080
          008080008080008080808080FFFFFF008080008080808080FFFFFF008080FFFF
          FF808080008080008080808080FFFFFF00808000808000808000808000808000
          80800000FF000080000080000080000080808080000080000080000080000080
          000080808080008080008080008080008080008080808080FFFFFF0080800080
          80008080808080FFFFFF808080008080008080008080008080808080FFFFFF00
          80800080800080800080800080800080800000FF000080000080000080000080
          0000800000800000800000808080800080800080800080800080800080800080
          80008080808080FFFFFF00808000808000808080808000808000808000808000
          8080FFFFFF808080008080008080008080008080008080008080008080008080
          0000FF0000800000800000800000800000800000808080800080800080800080
          80008080008080008080008080008080008080808080FFFFFF00808000808000
          8080008080008080008080FFFFFF808080008080008080008080008080008080
          0080800080800080800080800080800000800000800000800000800000808080
          8000808000808000808000808000808000808000808000808000808000808000
          8080808080FFFFFF008080008080008080008080008080808080008080008080
          0080800080800080800080800080800080800080800080800080800000FF0000
          8000008000008000008080808000808000808000808000808000808000808000
          8080008080008080008080008080008080808080FFFFFF008080008080008080
          8080800080800080800080800080800080800080800080800080800080800080
          800080800000FF00008000008000008000008000008080808000808000808000
          8080008080008080008080008080008080008080008080008080008080808080
          008080008080008080008080808080FFFFFF0080800080800080800080800080
          800080800080800080800080800000FF00008000008000008080808000008000
          0080000080808080008080008080008080008080008080008080008080008080
          008080008080808080008080008080008080008080008080808080FFFFFF0080
          800080800080800080800080800080800080800080800000FF00008000008000
          00808080800080800000FF000080000080000080808080008080008080008080
          008080008080008080008080008080808080008080008080008080808080FFFF
          FF008080008080808080FFFFFF00808000808000808000808000808000808000
          80800000FF0000800000808080800080800080800080800000FF000080000080
          000080808080008080008080008080008080008080008080808080FFFFFF0080
          80008080808080008080808080FFFFFF008080008080808080FFFFFF00808000
          80800080800080800080800080800080800000FF000080008080008080008080
          0080800080800000FF0000800000800000800080800080800080800080800080
          80008080808080FFFFFFFFFFFF808080008080008080008080808080FFFFFF00
          8080008080808080FFFFFF008080008080008080008080008080008080008080
          0080800080800080800080800080800080800080800000FF0000800000FF0080
          8000808000808000808000808000808000808080808080808000808000808000
          8080008080008080808080FFFFFFFFFFFFFFFFFF808080008080008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080808080808080808080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080008080008080008080
          008080008080008080008080008080008080}
        Layout = blGlyphTop
        ModalResult = 2
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
          03030606030303030303030303030303030303FFFF0303030303030303030303
          0303030303060404060303030303030303030303030303F8F8FF030303030303
          030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
          0303030303030303030303030303FE060603030303030303030303030303F8FF
          FFF8FF0303030303030303030303030303030303030303030303030303030303
          030303F8F8030303030303030303030303030303030304040603030303030303
          0303030303030303FFFF03030303030303030303030303030306060604030303
          0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
          0403030303030303030303030303F8FF03F8FF03030303030303030303030303
          03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
          030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
          0303030303030303030303FE060606040303030303030303030303F8FF0303F8
          FF030303030303030303030404030303FE060606040303030303030303FF0303
          F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
          F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
          030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
          06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
          060606060303030303030303F8FFFF030303030303F803030303030303030303
          0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
          0303030303030303030303030303030303030303030303F8F8F8F8F803030303
          0303}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        IsControl = True
      end
    end
  end
  object BZoomArticle: TBitBtn
    Tag = 100
    Left = 7
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Voir l'#39'article'
    Caption = 'Article'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    Visible = False
    OnClick = BZoomArticleClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomTiers: TBitBtn
    Tag = 100
    Left = 39
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Voir le client'
    Caption = 'Tiers'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    Visible = False
    OnClick = BZoomTiersClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomTarif: TBitBtn
    Tag = 100
    Left = 72
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Voir le tarif'
    Caption = 'Tarif'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    Visible = False
    OnClick = BZoomTarifClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomCommercial: TBitBtn
    Tag = 100
    Left = 136
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Voir le vendeur'
    Caption = 'Commercial'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    Visible = False
    OnClick = BZoomCommercialClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomPrecedente: TBitBtn
    Tag = 100
    Left = 169
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Pièce précédente'
    Caption = 'Précédente'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = BZoomPrecedenteClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomOrigine: TBitBtn
    Tag = 100
    Left = 201
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Pièce d'#39'origine'
    Caption = 'Origine'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Visible = False
    OnClick = BZoomOrigineClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object TDescriptif: TToolWindow97
    Left = 357
    Top = 103
    ClientHeight = 76
    ClientWidth = 253
    Caption = 'Descriptif détaillé'
    ClientAreaHeight = 76
    ClientAreaWidth = 253
    TabOrder = 7
    Visible = False
    OnClose = TDescriptifClose
    object Descriptif: THRichEditOLE
      Left = 0
      Top = 0
      Width = 253
      Height = 76
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
          'ial;}}'
        
          '{\*\generator Msftedit 5.41.15.1503;}\viewkind4\uc1\pard\f0\fs16' +
          ' '
        '\par '
        '\par }')
    end
  end
  object BZoomSuivante: TBitBtn
    Tag = 100
    Left = 265
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Pièce suivante'
    Caption = 'Suivante'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
    Visible = False
    OnClick = BZoomSuivanteClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomEcriture: TBitBtn
    Tag = 100
    Left = 233
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Ecriture comptable'
    Caption = 'Ecriture'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
    Visible = False
    OnClick = BZoomEcritureClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomDevise: TBitBtn
    Tag = 100
    Left = 298
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Voir la devise'
    Caption = 'Devise'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 10
    Visible = False
    OnClick = BZoomDeviseClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object PEntete: THPanel
    Left = 0
    Top = 0
    Width = 718
    Height = 86
    Align = alTop
    Color = clSilver
    FullRepaint = False
    TabOrder = 0
    OnExit = PEnteteExit
    OnResize = PEnteteResize
    BackGroundEffect = bdHorzIn
    ColorShadow = clWindowText
    ColorStart = clSilver
    TextEffect = tenone
    object FTitrePiece: THLabel
      Left = 29
      Top = 74
      Width = 25
      Height = 20
      Alignment = taCenter
      AutoSize = False
      Caption = 'TICKET'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      Transparent = True
      Visible = False
      OmbreColor = clNone
    end
    object HGP_DATEPIECE: THLabel
      Left = 228
      Top = 60
      Width = 15
      Height = 16
      Caption = 'du'
      FocusControl = GP_DATEPIECE
      Transparent = True
    end
    object HGP_TIERS: THLabel
      Left = 343
      Top = 60
      Width = 33
      Height = 16
      Caption = 'Client'
      FocusControl = GP_TIERS
      Transparent = True
    end
    object fAffTXT: TLCDLabel
      Left = 12
      Top = 5
      Width = 688
      Height = 48
      Caption = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
      PixelSize = pix4x4
      PixelSpacing = 1
      CharSpacing = 2
      LineSpacing = 0
      BorderSpace = 6
      TextLines = 1
      NoOfChars = 26
      BackGround = clBlack
      BorderColor = clYellow
      PixelOn = clLime
      PixelOff = clBlack
      UpSideDown = False
    end
    object GP_NUMEROPIECE: THLabel
      Left = 12
      Top = 58
      Width = 209
      Height = 20
      Alignment = taCenter
      AutoSize = False
      Caption = 'TICKET N° Non affecté'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
      Ombre = True
      OmbreDecalX = 1
      OmbreDecalY = 1
      OmbreColor = clGray
    end
    object HGP_REPRESENTANT: THLabel
      Left = 526
      Top = 60
      Width = 51
      Height = 16
      Caption = 'Vendeur'
      FocusControl = GP_REPRESENTANT
      Transparent = True
    end
    object GP_DATELIVRAISON: THCritMaskEdit
      Left = 579
      Top = 53
      Width = 26
      Height = 21
      Color = clYellow
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 9
      Text = '  /  /    '
      Visible = False
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object GP_DEPOT: THValComboBox
      Left = 495
      Top = 53
      Width = 25
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
      TabOrder = 4
      Visible = False
      OnChange = GP_DEPOTChange
      TagDispatch = 0
      DataType = 'GCDEPOT'
    end
    object GP_ETABLISSEMENT: THValComboBox
      Left = 467
      Top = 53
      Width = 25
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
      OnChange = GP_ETABLISSEMENTChange
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
    end
    object GP_REGIMETAXE: THValComboBox
      Left = 438
      Top = 53
      Width = 25
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
      TabOrder = 8
      Visible = False
      TagDispatch = 0
      DataType = 'TTREGIMETVA'
    end
    object GP_FACTUREHT: TCheckBox
      Left = 410
      Top = 53
      Width = 28
      Height = 17
      Caption = 'GP_FACTUREHT'
      Checked = True
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      State = cbChecked
      TabOrder = 7
      Visible = False
    end
    object GP_SAISIECONTRE: TCheckBox
      Left = 381
      Top = 53
      Width = 32
      Height = 17
      Caption = 'GP_SAISIECONTRE'
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
    object GP_DEVISE: THValComboBox
      Left = 60
      Top = 77
      Width = 25
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
      TabOrder = 5
      Visible = False
      OnChange = GP_DEVISEChange
      TagDispatch = 0
      DataType = 'TTDEVISE'
    end
    object GP_DATEPIECE: THCritMaskEdit
      Left = 249
      Top = 56
      Width = 78
      Height = 24
      Enabled = False
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 2
      Text = '  /  /    '
      OnExit = GP_DATEPIECEExit
      TagDispatch = 0
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object GP_TIERS: THCritMaskEdit
      Left = 382
      Top = 56
      Width = 136
      Height = 24
      CharCase = ecUpperCase
      TabOrder = 0
      OnDblClick = GP_TIERSDblClick
      OnEnter = GP_TIERSEnter
      OnExit = GP_TIERSExit
      TagDispatch = 0
      DataType = 'GCTIERSSAISIE'
      ElipsisButton = True
      OnElipsisClick = GP_TIERSElipsisClick
    end
    object GP_REPRESENTANT: THCritMaskEdit
      Left = 581
      Top = 56
      Width = 132
      Height = 24
      CharCase = ecUpperCase
      TabOrder = 1
      OnDblClick = GP_REPRESENTANTDblClick
      OnEnter = GP_REPRESENTANTEnter
      OnExit = GP_REPRESENTANTExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GP_REPRESENTANTElipsisClick
    end
  end
  object BZoomStock: TBitBtn
    Tag = 100
    Left = 330
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Voir le stock de l'#39'article'
    Caption = 'Stock'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 13
    Visible = False
    OnClick = BZoomStockClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BJustifTarif: TBitBtn
    Tag = 100
    Left = 104
    Top = 320
    Width = 28
    Height = 27
    Hint = 'Explication du tarif'
    Caption = 'Tarif'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 14
    Visible = False
    OnClick = BJustifTarifClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 256
    Top = 403
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 304
    Top = 403
  end
  object HTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Liste des articles'
      'Articles'
      'Tiers'
      'Liste des affaires'
      'Affaires'
      'ATTENTION : Pièce non enregistrée !'
      
        'ATTENTION : Cette pièce en cours de traitement par un autre util' +
        'isateur n'#39'a pas été enregistrée !'
      'Liste des commerciaux'
      'Conditionnements'
      'Conditionnement en cours :'
      'Non affecté'
      'Liste des commerciaux'
      'Liste des dépôts'
      
        'ATTENTION : Cette pièce ne peut pas passer en comptabilité et n'#39 +
        'a pas été enregistrée !'
      'Euro'
      'Autres'
      'Libellés automatiques'
      '(référence de substitution possible :'
      '(remplacement par référence :'
      'ATTENTION : l'#39'impression ne s'#39'est pas correctement effectuée !'
      'ATTENTION : la suppression ne s'#39'est pas correctement effectuée !'
      'Crédit accordé :'
      'Encours actuel :'
      
        'ATTENTION : La pièce présente un problème de numérotation et n'#39'a' +
        ' pas été enregistrée !'
      'Client'
      'Fournisseur'
      '26;'
      
        'ATTENTION. Le stock disponible est insuffisant pour certains art' +
        'icles.'
      'Changement de code TVA'
      '29;'
      'N°'
      'Liste des vendeurs'
      'Liste des motifs de démarque'
      'Choix du régime de taxe'
      '34;'
      
        'ATTENTION : la demande de détaxe ne s'#39'est pas correctement effec' +
        'tuée !'
      'Choix de l'#39'exception de taxation'
      'Choix du modèle de taxe'
      'ATTENTION : cet article ne correspond pas à ce fournisseur !')
    Left = 460
    Top = 403
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 104
    Top = 352
  end
  object HPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Il n'#39'existe pas de tarif pour cet article correspond' +
        'ant à la devise de la pièce !;E;O;O;O;'
      
        '1;?caption?;Saisie impossible. Il n'#39'existe pas de tarif pour cet' +
        ' article correspondant à la devise de la pièce !;E;O;O;O;'
      
        '2;?caption?;Saisie impossible. Cet article fermé n'#39'est pas autor' +
        'isé pour cette nature de document;E;O;O;O;'
      
        '3;?caption?;Choix impossible. Ce tiers fermé ne peut pas être ap' +
        'pelé pour cette nature de document;E;O;O;O;'
      
        '4;?caption?;Choix impossible. La nature du  tiers associé à l'#39'af' +
        'faire n'#39'est pas compatible avec la nature de pièce;E;O;O;O;'
      '5;?caption?;Vous devez renseigner un tiers valide;E;O;O;O;'
      '6;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;Y;N;'
      
        '7;?caption?;Voulez-vous affecter ce dépôt sur toutes les lignes ' +
        'concernées ?;Q;YN;Y;N;'
      
        '8;?caption?;Choix impossible. Ce code tiers est incorrect;E;O;O;' +
        'O;'
      
        '9;?caption?;Cette pièce ne peut pas être modifiée, seule la cons' +
        'ultation est autorisée;E;O;O;O;'
      
        '10;?caption?;Cet article est en rupture, confirmez-vous malgré t' +
        'out la quantité ?;Q;YN;Y;N;'
      '11;?caption?;Cet article est en rupture ;E;O;O;O;'
      '12;?caption?;Cet article n'#39'est plus géré ;E;O;O;O;'
      
        '13;?caption?;Le plafond de crédit accordé au tiers est dépassé !' +
        ' ;E;O;O;O;'
      '14;?caption?;Le crédit accordé au tiers est dépassé ! ;E;O;O;O;'
      
        '15;?caption?;Voulez-vous saisir ce taux dans la table de chancel' +
        'lerie ?;Q;YN;Y;N;O;'
      
        '16;?caption?;ATTENTION : Le taux en cours est de 1. Voulez-vous ' +
        'saisir ce taux dans la table de chancellerie;Q;YN;Y;N;O;'
      
        '17;?caption?;ATTENTION : la marge de la ligne est inférieure au ' +
        'minimum requis;E;O;O;O;'
      
        '18;?caption?;La date que vous avez renseignée n'#39'est pas valide;W' +
        ';O;O;O;'
      
        '19;?caption?;La date que vous avez renseignée n'#39'est pas dans un ' +
        'exercice ouvert;W;O;O;O;'
      
        '20;?caption?;La date que vous avez renseignée est antérieure à u' +
        'ne clôture;W;O;O;O;'
      
        '21;?caption?;La date que vous avez renseignée est antérieure à u' +
        'ne clôture;W;O;O;O;'
      
        '22;?caption?;La date que vous avez renseignée est en dehors des ' +
        'limites autorisées;W;O;O;O;'
      
        '23;?caption?;ATTENTION : la marge de la ligne est inférieure au ' +
        'minimum requis. Voulez-vous changer de code utilisateur ?;Q;YN;N' +
        ';N;'
      
        '24;?caption?;Le plafond de crédit accordé au tiers est dépassé. ' +
        'Voulez-vous changer de code utilisateur ? ;Q;YN;N;N;'
      '25;?caption?;Vous ne pouvez pas saisir avant le ;E;O;O;O;'
      
        '26;?caption?;Cette pièce contient déjà des lignes. Reprise des l' +
        'ignes de l'#39'affaire impossible ;E;O;O;O;'
      
        '27;?caption?;Attention : Vous n'#39'avez pas affecté d'#39'affaire sur c' +
        'ette pièce ;E;O;O;O;'
      
        '28;?caption?;La date est antérieure à celle de dernière clôture ' +
        'de stock;E;O;O;O;'
      
        '29;?caption?;ATTENTION. Cette opération est irréversible. Elle n' +
        'e réactivera pas la pièce d'#39'origine. Confirmez-vous la suppressi' +
        'on de la pièce ?;Q;YN;N;N;'
      
        '30;?caption?;Voulez-vous répercuter la date de livraison de l'#39'en' +
        'tête sur toutes les lignes?;Q;YN;Y;N;O;'
      
        '31;?caption?;Voulez-vous répercuter la date de livraison de l'#39'en' +
        'tête sur les lignes sélectionnées ?;Q;YN;Y;N;O;'
      '32;?caption?;Vous devez renseigner une devise;E;O;O;O;'
      '33;'
      '34;'
      '35;'
      '36;'
      '37;'
      
        '38;?caption?;Voulez-vous affecter cette TVA sur toutes les ligne' +
        's concernées ?;Q;YN;Y;N;'
      
        '39;?caption?;Voulez-vous affecter ce régime sur toute la pièce ?' +
        ';Q;YN;Y;N;'
      '40;'
      '41;?caption?;Cet article n'#39'est pas remisable;E;O;O;O;'
      '42;'
      '43;'
      '44;?caption?;Cet article n'#39'a pas de stock;E;O;O;O;'
      
        '45;?caption?;Saisie impossible. Le type de cet article est incom' +
        'patible avec les autres lignes de la pièce;E;O;O;O;'
      
        '46;?caption?;Saisie impossible. Cet article est interdit à la ve' +
        'nte.;E;O;O;O;'
      
        '47;?caption?;Vous avez modifié le prix d'#39'un article sans tarif, ' +
        'voulez-vous modifier le prix théorique ?;Q;YN;Y;N;'
      
        '48;?caption?;Le client est obligatoire pour ce motif de remise. ' +
        'Voulez-vous saisir le code du client ?;Q;YN;Y;N;'
      
        '49;?caption?;Le client est obligatoire pour cet article financie' +
        'r. Voulez-vous saisir le code du client ?;Q;YN;Y;N;'
      
        '50;?caption?;Le client est obligatoire pour ce montant de vente.' +
        ' Voulez-vous saisir le code du client ?;Q;YN;Y;N;'
      
        '51;?caption?;Saisie impossible. Le code du client est incompatib' +
        'le avec l'#39'article financier.;E;O;O;O;'
      
        '52;?caption?;Le client est obligatoire en cas de détaxe. Voulez-' +
        'vous saisir le code du client ?;Q;YN;Y;N;'
      '53;?caption?;ATTENTION. Les prix affichés sont en TTC.;W;O;O;O;'
      '54;?caption?;ATTENTION. Les prix affichés sont en HT.;W;O;O;O;'
      
        '55;?caption?;Vous êtes en dehors de la période d'#39'utilisation de ' +
        'ce motif de remise.;E;O;O;O;'
      
        '56;Fidélité;$$#10#10Voulez-vous faire bénéficier tout de suite d' +
        'e la fidélité à votre client ?;Q;YN;Y;N;'
      
        '57;?caption?;Le client est obligatoire pour livrer une vente. Vo' +
        'ulez-vous saisir le code du client ?;Q;YN;Y;N;'
      '58;'
      ''
      ''
      ''
      ''
      ''
      ''
      ''
      ''
      ''
      ' '
      ' ')
    Left = 501
    Top = 403
  end
  object PopD: TPopupMenu
    AutoPopup = False
    Left = 71
    Top = 352
    object POPD_S1: TMenuItem
      Caption = 'Francs'
      RadioItem = True
      OnClick = PopEuro
    end
    object POPD_S2: TMenuItem
      Caption = 'Euro'
      RadioItem = True
      OnClick = PopEuro
    end
    object POPD_S3: TMenuItem
      Caption = 'Autres'
      RadioItem = True
      OnClick = PopEuro
    end
  end
  object PopV: TPopupMenu
    AutoPopup = False
    Left = 136
    Top = 352
    object VPiece: TMenuItem
      Caption = 'Pièce'
      OnClick = VPieceClick
    end
    object VLigne: TMenuItem
      Caption = 'Ligne'
      OnClick = VLigneClick
    end
  end
  object PopL: TPopupMenu
    AutoPopup = False
    Left = 169
    Top = 352
    object TInsLigne: TMenuItem
      Caption = 'Insérer ligne'
      OnClick = TInsLigneClick
    end
    object TSupLigne: TMenuItem
      Caption = 'Supprimer ligne'
      OnClick = TSupLigneClick
    end
    object TCommentEnt: TMenuItem
      Caption = 'Commentaire d'#39'entête'
      OnClick = TCommentEntClick
    end
    object TCommentPied: TMenuItem
      Caption = 'Commentaire en pied'
      OnClick = TCommentPiedClick
    end
    object N4: TMenuItem
      Caption = '-'
    end
    object TRechArtImage: TMenuItem
      Caption = 'Recherche article par image'
      OnClick = TRechArtImageClick
    end
    object TLigneRetour: TMenuItem
      Caption = 'Ligne de retour'
      Enabled = False
      OnClick = TLigneRetourClick
    end
  end
  object HErr: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Erreurs en validation'
      
        '1;?caption?;Vous ne pouvez pas enregistrer une pièce vide;E;O;O;' +
        'O;'
      
        '2;?caption?;Vous ne pouvez pas enregistrer une pièce sans articl' +
        'e;E;O;O;O;'
      
        '3;?caption?;Vous ne pouvez pas enregistrer une pièce sans mouvem' +
        'ent;E;O;O;O;'
      
        '4;?caption?;Vous ne pouvez pas enregistrer une pièce à cette dat' +
        'e;E;O;O;O;'
      
        '5;?caption?;Enregistrement impossible : l'#39'acompte est supérieur ' +
        'au total de la pièce;E;O;O;O;'
      
        '6;?caption?;Enregistrement impossible : la devise est incorrecte' +
        ';E;O;O;O;'
      '7;'
      '8;'
      '9;'
      '10;'
      '11;'
      '12;'
      '13;'
      '14;'
      '15;'
      '16;'
      '17;'
      '18;'
      '19;'
      '20;'
      '21;?caption?;Le vendeur est obligatoire.;E;O;O;O;'
      
        '22;?caption?;En cas de remise vous devez indiquer un motif.;E;O;' +
        'O;O;'
      '23;?caption?;Le code du vendeur est incorrect.;E;O;O;O;'
      '24;?caption?;Le code du client est incorrect.;E;O;O;O;'
      
        '25;?caption?;Le code du client est incompatible avec une opérati' +
        'on de caisse.;E;O;O;O;'
      
        '26;?caption?;Le type de cet article est incompatible avec les au' +
        'tres lignes de la pièce;E;O;O;O;'
      
        '27;?caption?;La quantité de la ligne est supérieure au maximum a' +
        'utorisé;E;O;O;O;'
      
        '28;?caption?;Le prix unitaire de la ligne est supérieur au maxim' +
        'um autorisé;E;O;O;O;'
      
        '29;?caption?;Le règlement lié est aussi utilisé par une autre li' +
        'gne;E;O;O;O;'
      
        '30;?caption?;Annulation du ticket impossible car un règlement es' +
        't lié;E;O;O;O;'
      
        '31;?caption?;Vous avez dépassé le seuil de fidélité acquis par c' +
        'e client;E;O;O;O;'
      '')
    Left = 540
    Top = 403
  end
  object TimerGen: TTimer
    Enabled = False
    Interval = 0
    OnTimer = TimerGenTimer
    Left = 355
    Top = 403
  end
  object TimerLCD: TTimer
    Enabled = False
    Interval = 0
    OnTimer = TimerLCDTimer
    Left = 407
    Top = 403
  end
  object POPY: TPopupMenu
    AutoPopup = False
    Left = 6
    Top = 352
    object CpltEntete: TMenuItem
      Caption = 'Complément &Entête'
      OnClick = CpltEnteteClick
    end
    object Librepiece: TMenuItem
      Caption = '&Zones libres pièces'
      OnClick = LibrepieceClick
    end
    object CpltLigne: TMenuItem
      Caption = 'Complément &Ligne'
      OnClick = CpltLigneClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object AdrLiv: TMenuItem
      Caption = '&Adresses Livraison'
      Visible = False
      OnClick = AdrLivClick
    end
    object AdrFac: TMenuItem
      Caption = 'Adresses &Facturation'
      Visible = False
      OnClick = AdrFacClick
    end
    object N2: TMenuItem
      Caption = '-'
      Visible = False
    end
    object MBtarif: TMenuItem
      Caption = 'Mise à jour tarif'
      Visible = False
      OnClick = MBtarifClick
    end
    object MBDatesLivr: TMenuItem
      Caption = 'Mise à jour dates livraison lignes'
      Visible = False
    end
    object MBSoldeReliquat: TMenuItem
      Caption = 'Solder / Activer le reliquat'
      Visible = False
      OnClick = MBSoldeReliquatClick
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object MBDetailNomen: TMenuItem
      Caption = 'Détail nomenclature'
      OnClick = MBDetailNomenClick
    end
    object MBChangeDepot: TMenuItem
      Caption = 'Dépôt de la ligne'
      OnClick = MBChangeDepotClick
    end
  end
  object PopT: TPopupMenu
    AutoPopup = False
    Left = 39
    Top = 352
    object MCCreerClient: TMenuItem
      Caption = 'Création d'#39'un client'
      OnClick = MCCreerClientClick
    end
    object MCModifClient: TMenuItem
      Caption = 'Modification du client'
      OnClick = MCModifClientClick
    end
    object MCChoixClient: TMenuItem
      Caption = 'Choix du client'
      OnClick = MCChoixClientClick
    end
    object N7: TMenuItem
      Caption = '-'
    end
    object MCContactClient: TMenuItem
      Caption = 'Contacts du client'
      OnClick = MCContactClientClick
    end
    object N8: TMenuItem
      Caption = '-'
    end
    object MCFidelite: TMenuItem
      Caption = 'Cumul de fidélité'
      OnClick = MCFideliteClick
    end
    object MCHistoClient: TMenuItem
      Caption = 'Historique des achats'
      OnClick = MCHistoClientClick
    end
    object MCArticleClient: TMenuItem
      Caption = 'Liste des articles'
      OnClick = MCArticleClientClick
    end
    object MCSoldeClient: TMenuItem
      Caption = 'Solde du client'
      OnClick = MCSoldeClientClick
    end
  end
  object PopM: TPopupMenu
    AutoPopup = False
    Left = 201
    Top = 352
    object MDChoixModele: TMenuItem
      Caption = 'Modèle d'#39'impression'
      OnClick = MDChoixModeleClick
    end
    object MDReimpTicket: TMenuItem
      Caption = 'Réimpression du dernier ticket'
      OnClick = MDReimpTicketClick
    end
    object MDReimpBons: TMenuItem
      Caption = 'Réimpression des bons'
      OnClick = MDReimpBonsClick
    end
    object MDOuvreTiroir: TMenuItem
      Caption = 'Ouverture du tiroir caisse'
      OnClick = MDOuvreTiroirClick
    end
    object N5: TMenuItem
      Caption = '-'
    end
    object MDSituFlash: TMenuItem
      Caption = 'Situation Flash'
      OnClick = MDSituFlashClick
    end
    object MDStock: TMenuItem
      Caption = 'Consultation du stock'
      OnClick = MDStockClick
    end
    object N6: TMenuItem
      Caption = '-'
    end
    object MDDetaxe: TMenuItem
      Caption = 'Demande de détaxe'
      OnClick = MDDetaxeClick
    end
    object MDALivrer: TMenuItem
      Caption = 'Vente à livrer'
      OnClick = MDALivrerClick
    end
  end
  object PopX: TPopupMenu
    AutoPopup = False
    Left = 234
    Top = 352
    object MXExcepTaxe: TMenuItem
      Caption = 'Exception de taxation'
      OnClick = MXExcepTaxeClick
    end
    object MXTaxePiece: TMenuItem
      Caption = 'Taxation de la pièce'
      OnClick = MXTaxePieceClick
    end
    object MXTaxeLigne: TMenuItem
      Caption = 'Taxation de la ligne'
      OnClick = MXTaxeLigneClick
    end
  end
end

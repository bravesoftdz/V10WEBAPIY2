object FVentAna: TFVentAna
  Left = 468
  Top = 322
  HelpContext = 110000250
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSingle
  Caption = 'Ventilation par d'#233'faut de : '
  ClientHeight = 301
  ClientWidth = 428
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 236
    Top = 4
    Width = 38
    Height = 13
    Caption = 'TCache'
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    Visible = False
  end
  object TCache: THLabel
    Left = 164
    Top = 208
    Width = 3
    Height = 13
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object Label2: TLabel
    Left = 300
    Top = 84
    Width = 38
    Height = 13
    Caption = 'TCache'
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    Visible = False
  end
  object HPanel1: TPanel
    Left = 0
    Top = 266
    Width = 428
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object TFType: THLabel
      Left = 5
      Top = 11
      Width = 72
      Height = 13
      Caption = 'Ventilation type'
      FocusControl = FType
    end
    object FType: THValComboBox
      Left = 85
      Top = 7
      Width = 151
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnClick = FTypeClick
      TagDispatch = 0
      DataType = 'TTVENTILTYPE'
    end
    object Panel6: TPanel
      Left = 233
      Top = 2
      Width = 193
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object BSolde: THBitBtn
        Tag = 1
        Left = 66
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Solder '#224' 100 %'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BSoldeClick
        NumGlyphs = 2
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object OKBtn: THBitBtn
        Left = 99
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = OKBtnClick
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
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        IsControl = True
      end
      object BFermer: THBitBtn
        Left = 131
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BFermerClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object HelpBtn: THBitBtn
        Left = 163
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = HelpBtnClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 428
    Height = 239
    ActivePage = TS1
    Align = alClient
    TabOrder = 1
    OnChange = PagesChange
    object TS1: TTabSheet
      Caption = 'Axe 1'
      object FListe1: THGrid
        Left = 0
        Top = 0
        Width = 420
        Height = 211
        Align = alClient
        ColCount = 4
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        OnDblClick = FListe1DblClick
        OnKeyDown = FListe1KeyDown
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          '% Valeur;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnColExit = FListe1ColExit
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          92
          204
          76)
      end
    end
    object TS2: TTabSheet
      Caption = 'Axe 2'
      ImageIndex = 1
      object FListe2: THGrid
        Left = 0
        Top = 0
        Width = 420
        Height = 211
        Align = alClient
        ColCount = 4
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        OnDblClick = FListe1DblClick
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          '% Valeur;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          92
          204
          76)
      end
    end
    object TS3: TTabSheet
      Caption = 'Axe 3'
      ImageIndex = 2
      object FListe3: THGrid
        Left = 0
        Top = 0
        Width = 420
        Height = 211
        Align = alClient
        ColCount = 4
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        OnDblClick = FListe1DblClick
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          '% Valeur;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          92
          204
          76)
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 239
    Width = 428
    Height = 27
    Align = alBottom
    BevelOuter = bvNone
    Caption = ' '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object HLabel1: THLabel
      Left = 293
      Top = 7
      Width = 30
      Height = 13
      Caption = 'Total'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object FTotVal: THNumEdit
      Left = 339
      Top = 3
      Width = 60
      Height = 21
      Color = clBtnFace
      Ctl3D = True
      Enabled = False
      ParentCtl3D = False
      TabOrder = 0
      Decimals = 3
      Digits = 12
      Masks.PositiveMask = '# ##0.000'
      Debit = False
      UseRounding = True
      Validate = False
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Ventilations analytiques;Voulez-vous enregistrer les modificat' +
        'ions?;Q;YNC;Y;C;'
      
        '1;Ventilations analytiques;ATTENTION : La section est ferm'#233'e.;E;' +
        'O;O;O;'
      'Sections r'#233'ceptrices de :'
      'Ventilation par d'#233'faut de :'
      'Ventilation type :'
      'Ventilations analytiques pi'#232'ce'
      'Ventilations analytiques ligne'
      'Choix d'#39'une section analytique'
      'Ventilations anal. stock ent'#234'te '
      'Ventilations anal. stock ligne '
      
        '10;Ventilations analytiques;Attention ! La ventilation de l'#39'axe ' +
        '1 est nulle ou d'#233'passe 100 %. Confirmez-vous l'#39'enregistrement ?;' +
        'Q;YN;Y;Y;'
      
        '11;Ventilations analytiques;Attention ! La ventilation de l'#39'axe ' +
        '2 est nulle ou d'#233'passe 100 %. Confirmez-vous l'#39'enregistrement ?;' +
        'Q;YN;Y;Y;'
      
        '12;Ventilations analytiques;Attention ! La ventilation de l'#39'axe ' +
        '3 est nulle ou d'#233'passe 100 %. Confirmez-vous l'#39'enregistrement ?;' +
        'Q;YN;Y;Y;'
      ''
      '')
    Left = 212
    Top = 116
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'HPanel1'
      'BDelLigne'
      'BInsLigne'
      'BImprimer'
      'OkBtn'
      'BFermer'
      'HelpBtn')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 92
    Top = 112
  end
end

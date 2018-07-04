object FProfilRubrique: TFProfilRubrique
  Left = 79
  Top = 127
  Width = 637
  Height = 517
  HelpContext = 41110010
  BorderIcons = [biSystemMenu]
  Caption = 'Liste des rubriques du profil'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 449
    Width = 629
    Height = 34
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 30
      ClientWidth = 629
      Caption = 'ToolWindow971'
      ClientAreaHeight = 30
      ClientAreaWidth = 629
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      TabOrder = 0
      object BVALIDER: TToolbarButton97
        Left = 521
        Top = 1
        Width = 28
        Height = 27
        Flat = False
        NumGlyphs = 2
        OnClick = BVALIDERClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BFERMER: TToolbarButton97
        Left = 553
        Top = 1
        Width = 28
        Height = 27
        Flat = False
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BHELP: TToolbarButton97
        Left = 585
        Top = 1
        Width = 28
        Height = 27
        Flat = False
        OnClick = BHELPClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 489
        Top = 1
        Width = 28
        Height = 27
        Flat = False
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
  end
  object pages: TPageControl
    Left = 0
    Top = 0
    Width = 629
    Height = 449
    ActivePage = TSCritere
    Align = alClient
    MultiLine = True
    TabOrder = 1
    object TSCritere: TTabSheet
      Caption = 'Crit'#232're'
      object Splitter1: TSplitter
        Left = 321
        Top = 41
        Width = 4
        Height = 380
      end
      object GRID_PROFILRUB: THGrid
        Left = 325
        Top = 41
        Width = 296
        Height = 380
        Align = alClient
        ColCount = 4
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 0
        OnDblClick = BNRETIRE_RUBClick
        OnDragDrop = DEPOSE_OBJET
        OnDragOver = TEST_DEPOSE_OBJET
        OnMouseDown = GRID_PROFILRUBMouseDown
        SortedCol = 2
        Titres.Strings = (
          'Type'
          'N'#176
          'Rubrique'
          'E')
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
        ColWidths = (
          55
          40
          178
          14)
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 621
        Height = 41
        Align = alTop
        Caption = ' '
        TabOrder = 1
        object P: THLabel
          Left = 27
          Top = 12
          Width = 23
          Height = 13
          Caption = 'Profil'
        end
        object LProfil: TLabel
          Left = 56
          Top = 12
          Width = 413
          Height = 13
          AutoSize = False
          Caption = 'LProfil'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
      end
      object Panel2: TPanel
        Left = 0
        Top = 41
        Width = 321
        Height = 380
        Align = alLeft
        Caption = ' '
        TabOrder = 2
        object BNONRECUP_REM: THSpeedButton
          Left = 289
          Top = 120
          Width = 23
          Height = 22
          NumGlyphs = 2
          OnClick = BNONRECUP_REMClick
          GlobalIndexImage = 'Z0056_S16G2'
        end
        object BNONRECUP_COT: THSpeedButton
          Left = 289
          Top = 308
          Width = 23
          Height = 22
          NumGlyphs = 2
          OnClick = BNONRECUP_COTClick
          GlobalIndexImage = 'Z0056_S16G2'
        end
        object BNRETIRE_RUB: THSpeedButton
          Left = 288
          Top = 188
          Width = 23
          Height = 22
          NumGlyphs = 2
          OnClick = BNRETIRE_RUBClick
          GlobalIndexImage = 'Z0077_S16G2'
        end
        object THEMEREM: TLabel
          Left = 11
          Top = 12
          Width = 33
          Height = 13
          Caption = 'Th'#232'me'
        end
        object THEMECOT: TLabel
          Left = 12
          Top = 196
          Width = 33
          Height = 13
          Caption = 'Th'#232'me'
        end
        object GRID_REMUNERATION: THGrid
          Left = 9
          Top = 32
          Width = 272
          Height = 150
          ColCount = 2
          DefaultRowHeight = 18
          FixedCols = 0
          RowCount = 2
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
          TabOrder = 0
          OnDblClick = BNONRECUP_REMClick
          OnDragDrop = DEPOSE_OBJET
          OnDragOver = TEST_DEPOSE_OBJET
          OnMouseDown = GRID_PROFILRUBMouseDown
          SortedCol = 0
          Titres.Strings = (
            'N'#176
            'R'#233'mun'#233'ration')
          Couleur = False
          MultiSelect = True
          TitleBold = True
          TitleCenter = True
          ColCombo = 0
          SortEnabled = True
          SortRowExclude = 0
          TwoColors = False
          AlternateColor = 13224395
          ColWidths = (
            55
            209)
        end
        object GRID_COTISATION: THGrid
          Left = 9
          Top = 219
          Width = 272
          Height = 150
          ColCount = 2
          DefaultRowHeight = 18
          FixedCols = 0
          RowCount = 2
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
          TabOrder = 1
          OnDblClick = BNONRECUP_COTClick
          OnDragDrop = DEPOSE_OBJET
          OnDragOver = TEST_DEPOSE_OBJET
          OnMouseDown = GRID_PROFILRUBMouseDown
          SortedCol = 1
          Titres.Strings = (
            'N'#176
            'Cotisation')
          Couleur = False
          MultiSelect = True
          TitleBold = True
          TitleCenter = True
          ColCombo = 0
          SortEnabled = True
          SortRowExclude = 0
          TwoColors = False
          AlternateColor = 13224395
          ColWidths = (
            55
            211)
        end
        object CbxThemeRem: THValComboBox
          Left = 56
          Top = 8
          Width = 225
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 2
          OnChange = CbxThemeRemChange
          TagDispatch = 0
          Vide = True
          DataType = 'PGTHEMEREM'
        end
        object CbxThemeCot: THValComboBox
          Left = 56
          Top = 192
          Width = 225
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 3
          OnChange = CbxThemeCotChange
          TagDispatch = 0
          Vide = True
          DataType = 'PGTHEMECOT'
        end
      end
      object FLISTE: TCheckBox
        Left = 428
        Top = 12
        Width = 189
        Height = 17
        Caption = 'Liste d'#39'exportation (non visible)'
        TabOrder = 3
        Visible = False
      end
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous sauvegarder les modifications ?;Q;YNC;Y;' +
        'C')
    Left = 296
    Top = 31
  end
end

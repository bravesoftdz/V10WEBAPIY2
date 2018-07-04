object FGenere: TFGenere
  Left = 361
  Top = 161
  Width = 434
  Height = 389
  Caption = 'G'#233'n'#233'ration des '#233'critures pour les affaires en cours'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 426
    Height = 286
    ActivePage = Extraction
    Align = alClient
    ParentShowHint = False
    ShowHint = False
    TabOrder = 0
    object Extraction: TTabSheet
      Caption = 'Extraction'
      object HS_SECTION: THLabel
        Left = 36
        Top = 64
        Width = 56
        Height = 13
        Caption = '&Sections de'
      end
      object HS_SECTIONJOKER: THLabel
        Left = 36
        Top = 64
        Width = 41
        Height = 13
        Caption = '&Sections'
        Visible = False
      end
      object HS_SECTION_: TLabel
        Left = 262
        Top = 64
        Width = 6
        Height = 13
        Caption = #224
      end
      object TS_AXE: THLabel
        Left = 36
        Top = 32
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = S_AXE
      end
      object HLabel1: THLabel
        Left = 36
        Top = 96
        Width = 99
        Height = 13
        AutoSize = False
        Caption = '&Sections d'#39'exception'
        FocusControl = FExcep
      end
      object HLabel4: THLabel
        Left = 36
        Top = 128
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = FExercice
      end
      object HLabel6: THLabel
        Left = 36
        Top = 160
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Dates comptables du'
        FocusControl = FDateExtract1
      end
      object Label7: TLabel
        Left = 260
        Top = 160
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = FDateExtract2
      end
      object HLabel10: THLabel
        Left = 36
        Top = 192
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Type d'#39#233'criture'
        FocusControl = FDateExtract1
      end
      object S_AXE: THValComboBox
        Left = 167
        Top = 28
        Width = 201
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnChange = S_AXEChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object FExcep: TEdit
        Left = 167
        Top = 92
        Width = 201
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 1
      end
      object FExercice: THValComboBox
        Left = 167
        Top = 124
        Width = 201
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 2
        OnChange = FExerciceChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object FDateExtract1: TMaskEdit
        Left = 167
        Top = 156
        Width = 70
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '  /  /    '
        OnKeyPress = FDateExtract1KeyPress
      end
      object FDateExtract2: TMaskEdit
        Left = 296
        Top = 156
        Width = 70
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = FDateExtract1KeyPress
      end
      object FQualifPiece: THMultiValComboBox
        Left = 167
        Top = 188
        Width = 201
        Height = 21
        TabOrder = 5
        Text = '<<Tous>>'
        Abrege = False
        DataType = 'TTQUALPIECE'
        Complete = False
        OuInclusif = False
      end
      object S_SECTION: THCritMaskEdit
        Left = 167
        Top = 60
        Width = 89
        Height = 21
        TabOrder = 6
        OnChange = S_SECTIONChange
        OnExit = S_SECTIONExit
        TagDispatch = 0
        Plus = 'S_AFFAIREENCOURS="X"'
        DataType = 'TZSECTIONTOUS'
        Operateur = Superieur
        ElipsisButton = True
      end
      object S_SECTION_: THCritMaskEdit
        Left = 279
        Top = 60
        Width = 89
        Height = 21
        TabOrder = 7
        OnExit = S_SECTIONExit
        TagDispatch = 0
        Plus = 'S_AFFAIREENCOURS="X"'
        DataType = 'TZSECTIONTOUS'
        Operateur = Inferieur
        ElipsisButton = True
      end
    end
    object Pzlibre: TTabSheet
      Caption = 'Tables Libres Sections'
      ImageIndex = 2
      object TS_TABLE0: TLabel
        Left = 4
        Top = 18
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE0'
        FocusControl = S_TABLE0
      end
      object TS_TABLE1: TLabel
        Left = 4
        Top = 65
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE1'
        FocusControl = S_TABLE1
      end
      object TS_TABLE2: TLabel
        Left = 4
        Top = 112
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE2'
        FocusControl = S_TABLE2
      end
      object TS_TABLE3: TLabel
        Left = 4
        Top = 159
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE3'
        FocusControl = S_TABLE3
      end
      object TS_TABLE4: TLabel
        Left = 4
        Top = 206
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE4'
        FocusControl = S_TABLE4
      end
      object TS_TABLE5: TLabel
        Left = 213
        Top = 18
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE5'
        FocusControl = S_TABLE5
      end
      object TS_TABLE6: TLabel
        Left = 213
        Top = 65
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE6'
        FocusControl = S_TABLE6
      end
      object TS_TABLE7: TLabel
        Left = 213
        Top = 112
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE7'
        FocusControl = S_TABLE7
      end
      object TS_TABLE8: TLabel
        Left = 213
        Top = 159
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE8'
        FocusControl = S_TABLE8
      end
      object TS_TABLE9: TLabel
        Left = 213
        Top = 206
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TS_TABLE9'
        FocusControl = S_TABLE9
      end
      object S_TABLE0: THCpteEdit
        Left = 104
        Top = 14
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatSect0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE1: THCpteEdit
        Left = 104
        Top = 61
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatSect1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE2: THCpteEdit
        Left = 104
        Top = 108
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatSect2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE3: THCpteEdit
        Left = 104
        Top = 155
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatSect3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE4: THCpteEdit
        Left = 104
        Top = 202
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatSect4
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE5: THCpteEdit
        Left = 314
        Top = 14
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatSect5
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE6: THCpteEdit
        Left = 314
        Top = 61
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatSect6
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE7: THCpteEdit
        Left = 314
        Top = 108
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatSect7
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE8: THCpteEdit
        Left = 314
        Top = 155
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatSect8
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object S_TABLE9: THCpteEdit
        Left = 314
        Top = 202
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        ZoomTable = tzNatSect9
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    object GenTlibre: TTabSheet
      Caption = 'Tables libres G'#233'n'#233'raux'
      ImageIndex = 3
      object TG_TABLE0: TLabel
        Left = 4
        Top = 18
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE0'
        FocusControl = G_TABLE0
      end
      object TG_TABLE1: TLabel
        Left = 4
        Top = 65
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE1'
        FocusControl = G_TABLE1
      end
      object TG_TABLE2: TLabel
        Left = 4
        Top = 106
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE2'
        FocusControl = G_TABLE2
      end
      object TG_TABLE3: TLabel
        Left = 4
        Top = 153
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE3'
        FocusControl = G_TABLE3
      end
      object TG_TABLE4: TLabel
        Left = 4
        Top = 201
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE4'
        FocusControl = G_TABLE4
      end
      object TG_TABLE5: TLabel
        Left = 212
        Top = 18
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE5'
        FocusControl = G_TABLE5
      end
      object TG_TABLE6: TLabel
        Left = 212
        Top = 65
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE6'
        FocusControl = G_TABLE6
      end
      object TG_TABLE7: TLabel
        Left = 212
        Top = 106
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE7'
        FocusControl = G_TABLE7
      end
      object TG_TABLE8: TLabel
        Left = 212
        Top = 153
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE8'
        FocusControl = G_TABLE8
      end
      object TG_TABLE9: TLabel
        Left = 212
        Top = 201
        Width = 98
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE9'
        FocusControl = G_TABLE9
      end
      object G_TABLE0: THCpteEdit
        Left = 104
        Top = 14
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatGene0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE1: THCpteEdit
        Left = 104
        Top = 61
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE2: THCpteEdit
        Left = 104
        Top = 102
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatGene0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE3: THCpteEdit
        Left = 104
        Top = 149
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE4: THCpteEdit
        Left = 104
        Top = 197
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE5: THCpteEdit
        Left = 312
        Top = 14
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatGene0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE6: THCpteEdit
        Left = 312
        Top = 61
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE7: THCpteEdit
        Left = 312
        Top = 102
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatGene0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE8: THCpteEdit
        Left = 312
        Top = 149
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE9: THCpteEdit
        Left = 312
        Top = 197
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    object Generation: TTabSheet
      Caption = 'G'#233'n'#233'ration'
      ImageIndex = 1
      object HLabel2: THLabel
        Left = 229
        Top = 21
        Width = 91
        Height = 13
        Caption = '&Date de g'#233'n'#233'ration'
        FocusControl = FDateGenere
      end
      object TFGen: THLabel
        Left = 22
        Top = 21
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Journal de g'#233'n'#233'ration'
      end
      object TSelectCpte: THLabel
        Left = 22
        Top = 53
        Width = 106
        Height = 13
        AutoSize = False
        Caption = '&Type d'#39#233'criture'
        Color = clBtnFace
        FocusControl = FQualGenere
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object LFile: THLabel
        Left = 22
        Top = 85
        Width = 68
        Height = 13
        Caption = 'Nom du fichier'
        FocusControl = FileName
        Visible = False
      end
      object RechFile: TToolbarButton97
        Left = 378
        Top = 80
        Width = 17
        Height = 23
        Hint = 'Parcourir'
        Caption = '...'
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = RechFileClick
      end
      object FDateGenere: TMaskEdit
        Left = 326
        Top = 17
        Width = 70
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '  /  /    '
        OnKeyPress = FDateExtract1KeyPress
      end
      object FJalGenere: THCpteEdit
        Left = 142
        Top = 17
        Width = 70
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzJOD
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FQualGenere: THValComboBox
        Left = 142
        Top = 49
        Width = 253
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
      object FileName: TEdit
        Left = 142
        Top = 81
        Width = 234
        Height = 21
        TabOrder = 3
        Visible = False
      end
      object GroupBox1: TGroupBox
        Left = 8
        Top = 128
        Width = 201
        Height = 118
        Caption = 'Charges '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object HLabel3: THLabel
          Left = 6
          Top = 28
          Width = 76
          Height = 13
          Caption = 'Compte de bilan'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel5: THLabel
          Left = 6
          Top = 60
          Width = 88
          Height = 13
          Caption = 'Compte de gestion'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TG_TABLE11: THLabel
          Left = 6
          Top = 92
          Width = 43
          Height = 13
          Caption = 'Rubrique'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object FCptCB: THCpteEdit
          Left = 104
          Top = 24
          Width = 76
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          ZoomTable = tzGNonCollectif
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
        object FCptCG: THCpteEdit
          Left = 104
          Top = 56
          Width = 76
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          ZoomTable = tzGNonCollectif
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
        object G_TABLEC: THCpteEdit
          Left = 104
          Top = 87
          Width = 76
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          ZoomTable = tzRubCPTA
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
      end
      object GroupBox2: TGroupBox
        Left = 215
        Top = 128
        Width = 194
        Height = 118
        Caption = 'Produit '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        object HLabel7: THLabel
          Left = 8
          Top = 28
          Width = 76
          Height = 13
          Caption = 'Compte de bilan'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel8: THLabel
          Left = 6
          Top = 58
          Width = 88
          Height = 13
          Caption = 'Compte de gestion'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel9: THLabel
          Left = 6
          Top = 91
          Width = 43
          Height = 13
          Caption = 'Rubrique'
          FocusControl = G_TABLEP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object FCptPB: THCpteEdit
          Left = 104
          Top = 24
          Width = 76
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          ZoomTable = tzGNonCollectif
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
        object FCptPG: THCpteEdit
          Left = 104
          Top = 54
          Width = 76
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          ZoomTable = tzGNonCollectif
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
        object G_TABLEP: THCpteEdit
          Left = 104
          Top = 87
          Width = 76
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          ZoomTable = tzNatGene0
          Vide = False
          Bourre = False
          okLocate = True
          SynJoker = False
        end
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 286
    Width = 426
    Height = 76
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 40
      ClientHeight = 32
      ClientWidth = 426
      Caption = 'Barre d'#39'outils'
      ClientAreaHeight = 32
      ClientAreaWidth = 426
      DockPos = 0
      DockRow = 1
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 327
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Lancer l'#39#233'tat'
        DisplayMode = dmGlyphOnly
        Caption = 'OK'
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
        Left = 359
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 391
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
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object LEnCours: TLabel
        Left = 8
        Top = 10
        Width = 46
        Height = 13
        Caption = 'LEnCours'
        Visible = False
      end
    end
    object PanelFiltre: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 426
      Caption = 'Filtres'
      ClientAreaHeight = 36
      ClientAreaWidth = 426
      DockPos = 0
      FullSize = True
      TabOrder = 1
      object BFiltre: TToolbarButton97
        Left = 7
        Top = 7
        Width = 69
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
      end
      object FFiltres: THValComboBox
        Left = 80
        Top = 8
        Width = 313
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object Sauve: TSaveDialog
    Filter = 
      'Fichiers textes (*.txt).|*.txt|Fichiers ASCII (*.asc).|*.asc|Fic' +
      'hiers SAARI (*.pn*).|*.pn*|Fichiers EDIFICAS (*.edf).|*.edf|Tous' +
      ' fichiers (*.*).|*.*'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofNoLongNames]
    Title = 'Exportation des mouvements'
    Left = 91
    Top = 319
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    OnBeforeTraduc = HMTradBeforeTraduc
    Traduction = True
    Left = 145
    Top = 319
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Plan des affaires en cours'
      '1;?caption?;Erreur : Le journal n'#39'existe pas;E;O;O;O;'
      
        '2;?caption?;Erreur : Le compte de bilan pour g'#233'n'#233'ration des char' +
        'ges n'#39'existe pas;E;O;O;O;'
      
        '3;?caption?;Erreur : Le compte de gestion pour g'#233'n'#233'ration des ch' +
        'arges n'#39'existe pas;E;O;O;O;'
      
        '4;?caption?;Erreur : Le compte de bilan pour g'#233'n'#233'ration des prod' +
        'uits n'#39'existe pas;E;O;O;O;'
      
        '5;?caption?;Erreur : Le compte de gestion pour g'#233'n'#233'ration des pr' +
        'oduits n'#39'existe pas;E;O;O;O;'
      
        '6;?caption?;Erreur : Le journal n'#39'a pas de souche comptable;E;O;' +
        'O;O;'
      
        '7;?caption?;Erreur : Le journal n'#39'a pas de souche de simulation;' +
        'E;O;O;O;'
      
        '8;?caption?;Erreur : La table libre des compte g'#233'n'#233'raux doit '#234'tr' +
        'e renseign'#233'e;E;O;O;O;'
      
        '9;?caption?;Erreur : La rubrique utilis'#233'e pour les charges doit ' +
        #234'tre renseign'#233'e;E;O;O;O;'
      
        '10;?caption?;Erreur : La rubrique utilis'#233'e pour les produits doi' +
        't '#234'tre renseign'#233'e;E;O;O;O;'
      
        '11;?caption?;Confirmez-vous la g'#233'n'#233'ration des '#233'critures ?;Q;YN;Y' +
        ';Y;'
      '12;?caption?;Le traitement s'#39'est correctement d'#233'roul'#233';E;O;O;O;'
      'Traitement en cours de l'#39'affaire '
      
        '14;?caption?;Vous n'#39'avez pas demand'#233' la constitution d'#39'un fichie' +
        'r ASCII. Confirmez-vous ?;Q;YN;Y;Y;'
      
        '15;?caption?;Erreur : Le compte de bilan pour g'#233'n'#233'ration des cha' +
        'rges est ventilable;E;O;O;O;'
      
        '16;?caption?;Erreur : Le compte de gestion pour g'#233'n'#233'ration des c' +
        'harges n'#39'est pas ventilable;E;O;O;O;'
      
        '17;?caption?;Erreur : Le compte de bilan pour g'#233'n'#233'ration des pro' +
        'duits est ventilable;E;O;O;O;'
      
        '18;?caption?;Erreur : Le compte de gestion pour g'#233'n'#233'ration des p' +
        'roduits n'#39'est pas ventilable;E;O;O;O;'
      '19;?caption?;Aucun mouvement n'#39'a pu '#234'tre g'#233'n'#233'r'#233';E;O;O;O;'
      'Constitution du fichier au format PGI'
      
        '21;?caption?;Erreur : La date de g'#233'n'#233'ration n'#39'est pas correcte.;' +
        'E;O;O;O;'
      ''
      ' '
      ' ')
    Left = 180
    Top = 319
  end
  object POPF: TPopupMenu
    Left = 220
    Top = 320
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

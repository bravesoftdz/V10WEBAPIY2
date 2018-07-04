object ExtPConcept: TExtPConcept
  Left = 137
  Top = 218
  Width = 858
  Height = 480
  HelpContext = 1040
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Param'#233'trage'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 149
    Top = 0
    Height = 410
  end
  object Panel1: TPanel
    Left = 0
    Top = 410
    Width = 842
    Height = 34
    Align = alBottom
    BevelOuter = bvLowered
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
    DesignSize = (
      842
      34)
    object Binsert: TToolbarButton97
      Left = 61
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Cr'#233'ation champ fin de liste'
      DisplayMode = dmGlyphOnly
      Caption = 'Nouveau'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = Champnormal1Click
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object BDelete: TToolbarButton97
      Left = 90
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Supprimer'
      DisplayMode = dmGlyphOnly
      Caption = 'Supprimer'
      ParentShowHint = False
      ShowHint = True
      OnClick = Supprimer1Click
      GlobalIndexImage = 'Z0005_S16G1'
    end
    object bCopier: TToolbarButton97
      Left = 119
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Copier'
      AllowAllUp = True
      OnClick = Copier1Click
      GlobalIndexImage = 'Z0342_S16G1'
    end
    object BColler: TToolbarButton97
      Left = 148
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Coller'
      AllowAllUp = True
      OnClick = Coller1Click
      GlobalIndexImage = 'Z0175_S16G1'
    end
    object Bcouper: TToolbarButton97
      Left = 177
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Couper'
      AllowAllUp = True
      OnClick = Couper1Click
      GlobalIndexImage = 'Z1067_S16G1'
    end
    object Bapprecu: TToolbarButton97
      Left = 206
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Scruter'
      AllowAllUp = True
      OnClick = Scruter1Click
      GlobalIndexImage = 'Z0138_S16G1'
    end
    object BImprimer: TToolbarButton97
      Left = 721
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Anchors = [akTop, akRight]
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
      Left = 753
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      Anchors = [akTop, akRight]
      NumGlyphs = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = BValiderClick
      GlobalIndexImage = 'Z0003_S16G2'
    end
    object BAnnuler: TToolbarButton97
      Left = 785
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Anchors = [akTop, akRight]
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      OnClick = BAnnulerClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: TToolbarButton97
      Tag = 1040
      Left = 817
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object bZoomPiece: TToolbarButton97
      Left = 689
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Execution du Script'
      HelpContext = 1710
      Anchors = [akTop, akRight]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 2
      ParentFont = False
      Spacing = -1
      OnClick = bZoomPieceClick
      GlobalIndexImage = 'Z0016_S16G1'
      IsControl = True
    end
    object bSelectAll: TToolbarButton97
      Left = 32
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Insertion champ'
      HelpContext = 1710
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 3
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = miNouveauClick
      GlobalIndexImage = 'Z0291_S16G1'
      IsControl = True
    end
    object bDown: TToolbarButton97
      Left = 3
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Affectation des variables'
      HelpContext = 1710
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 5
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = bDownClick
      GlobalIndexImage = 'Z0548_S16G1'
      IsControl = True
    end
    object BColor: TPaletteButton97
      Left = 235
      Top = 4
      Width = 36
      Height = 27
      Hint = 'Couleur'
      DisplayMode = dmGlyphOnly
      DropdownAlways = True
      DropdownArrow = True
      DropdownCombo = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = [fsBold]
      Glyph.Data = {
        66010000424D6601000000000000760000002800000014000000140000000100
        040000000000F000000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
        7777777700007777777777777777777700007777777777777777777700007777
        7777777777777777000077777777777777777777000077777777777777777777
        0000777777777777777777770000774467704448777777770000774446770F44
        4777777700007777777770F44477777700007777777770F44447777700007777
        7777770F44447777000077777777770FF44477770000777777777770FF447777
        000077777777777700007777000077777777777780BB07770000777777777777
        70BB07770000777777777777770BB07700007777777777777777777700007777
        77777777777777770000}
      GlyphMask.Data = {00000000}
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      CurrentChoix = 0
      SecondChoix = 0
      PaletteType = cbColor
      OnChange = BColorChange
      Device = fdScreen
    end
    object Breprise: TToolbarButton97
      Left = 270
      Top = 4
      Width = 33
      Height = 27
      Hint = 'Reprise'
      HelpContext = 1710
      DropdownArrow = True
      DropdownMenu = POPZ
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 3
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      GlobalIndexImage = 'Z1334_S16G1'
      IsControl = True
    end
    object BRecherche: TToolbarButton97
      Left = 304
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Rechercher'
      AllowAllUp = True
      OnClick = BRechercheClick
      GlobalIndexImage = 'Z0077_S16G1'
    end
    object bOptions: TToolbarButton97
      Left = 333
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Options'
      AllowAllUp = True
      Visible = False
      OnClick = bOptionsClick
      GlobalIndexImage = 'Z0335_S16G1'
    end
  end
  object Panel2: TPanel
    Left = 152
    Top = 0
    Width = 690
    Height = 410
    Align = alClient
    TabOrder = 1
    object Memo1: TASCIIView
      Left = 1
      Top = 225
      Width = 688
      Height = 144
      Align = alClient
      BevelInner = bvNone
      BevelOuter = bvRaised
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Fixedsys'
      Font.Style = []
      Margin = True
      TabOrder = 0
      TabStop = True
      OnGetLine = Memo1GetLine
      OnMouseDown = Memo1MouseDown
      OnMouseMove = Memo1MouseMove
      OnCaretMove = Memo1CaretMove
    end
    object Panel3: TPanel
      Left = 1
      Top = 369
      Width = 688
      Height = 40
      Align = alBottom
      TabOrder = 1
      DesignSize = (
        688
        40)
      object Label1: TLabel
        Left = 8
        Top = 13
        Width = 29
        Height = 13
        Caption = 'D'#233'but'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label2: TLabel
        Left = 70
        Top = 13
        Width = 45
        Height = 13
        Caption = 'Longueur'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label3: TLabel
        Left = 150
        Top = 13
        Width = 14
        Height = 13
        Caption = 'Fin'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Lbdeb: TLabel
        Left = 49
        Top = 13
        Width = 3
        Height = 13
      end
      object lbLon: TLabel
        Left = 122
        Top = 13
        Width = 3
        Height = 13
      end
      object lbFin: TLabel
        Left = 183
        Top = 13
        Width = 3
        Height = 13
      end
      object btnAffecter: TToolbarButton97
        Left = 343
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Affectation'
        HelpContext = 1710
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 5
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = btnAffecterClick
        GlobalIndexImage = 'Z0118_S16G1'
        IsControl = True
      end
      object lblChampCourant: TLabel
        Left = 589
        Top = 20
        Width = 3
        Height = 13
        Anchors = [akTop, akRight]
      end
      object lblEnregCourant: TLabel
        Left = 564
        Top = 4
        Width = 22
        Height = 13
        Anchors = [akTop, akRight]
        Caption = 'TRA'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object btnArboresCheminCourant: TToolbarButton97
        Left = 569
        Top = 19
        Width = 17
        Height = 17
        Anchors = [akTop, akRight]
        Enabled = False
        NoBorder = True
        Visible = False
        GlobalIndexImage = 'Z0235_S16G1'
      end
      object Label40: TLabel
        Left = 517
        Top = 4
        Width = 43
        Height = 13
        Anchors = [akTop, akRight]
        Caption = 'Position :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HPanel1: THPanel
        Left = 189
        Top = 7
        Width = 138
        Height = 25
        BevelOuter = bvLowered
        Color = clBlack
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object Label14: TLabel
          Left = 4
          Top = 4
          Width = 108
          Height = 18
          AutoSize = False
          Caption = 'C:1 L:1'
          Color = clBlack
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGreen
          Font.Height = -17
          Font.Name = 'Small Fonts'
          Font.Style = [fsBold]
          ParentColor = False
          ParentFont = False
        end
        object Label15: TLabel
          Left = 6
          Top = 2
          Width = 115
          Height = 17
          AutoSize = False
          Caption = 'C:1 L:1'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clLime
          Font.Height = -17
          Font.Name = 'Small Fonts'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
      end
    end
    object PageControl1: TPageControl
      Left = 1
      Top = 1
      Width = 688
      Height = 224
      ActivePage = tsComplement
      Align = alTop
      MultiLine = True
      TabOrder = 2
      OnChange = PageControl1Change
      object tbDEFINITION: TTabSheet
        Caption = 'D'#233'finition'
        object Label4: TLabel
          Left = 62
          Top = 12
          Width = 22
          Height = 13
          Caption = 'Nom'
        end
        object Label5: TLabel
          Left = 62
          Top = 38
          Width = 45
          Height = 13
          Caption = 'Longueur'
        end
        object Label6: TLabel
          Left = 260
          Top = 11
          Width = 24
          Height = 13
          Caption = 'Type'
        end
        object Label23: TLabel
          Left = 402
          Top = 75
          Width = 74
          Height = 13
          Caption = 'Type de champ'
        end
        object Label13: TLabel
          Left = 402
          Top = 106
          Width = 88
          Height = 13
          Caption = 'Nbre de d'#233'cimales'
        end
        object Label21: TLabel
          Left = 257
          Top = 120
          Width = 148
          Height = 13
          Caption = 'Mise '#224' jour du libelle de compte'
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object Label20: TLabel
          Left = 257
          Top = 146
          Width = 115
          Height = 13
          Caption = 'Op'#233'ration sur champ nul'
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object Label27: TLabel
          Left = 220
          Top = 38
          Width = 64
          Height = 13
          Caption = 'Date d'#39'entr'#233'e'
        end
        object Label29: TLabel
          Left = 404
          Top = 38
          Width = 68
          Height = 13
          Caption = 'Date de Sortie'
        end
        object GroupBox1: TGroupBox
          Left = 54
          Top = 56
          Width = 345
          Height = 49
          BiDiMode = bdLeftToRight
          Caption = ' Positionnement '
          ParentBiDiMode = False
          TabOrder = 0
          object lblDebut: TLabel
            Left = 8
            Top = 20
            Width = 29
            Height = 13
            Caption = 'D'#233'but'
          end
          object Label7: TLabel
            Left = 216
            Top = 20
            Width = 45
            Height = 13
            Caption = 'Longueur'
          end
          object lblFin: TLabel
            Left = 120
            Top = 20
            Width = 14
            Height = 13
            Caption = 'Fin'
          end
          object ChampDebut: TSpinEdit
            Left = 56
            Top = 15
            Width = 57
            Height = 23
            AutoSize = False
            MaxValue = 4096
            MinValue = 1
            TabOrder = 0
            Value = 1
            OnChange = ChampDebutChange
          end
          object ChampLon: TSpinEdit
            Left = 280
            Top = 15
            Width = 56
            Height = 23
            AutoSize = False
            MaxValue = 255
            MinValue = 0
            TabOrder = 1
            Value = 0
            OnChange = ChampLonChange
          end
          object edFin: TSpinEdit
            Left = 152
            Top = 15
            Width = 56
            Height = 22
            MaxValue = 4096
            MinValue = 1
            TabOrder = 2
            Value = 1
            OnChange = edFinChange
          end
        end
        object edLongueur: TSpinEdit
          Left = 118
          Top = 33
          Width = 73
          Height = 23
          AutoSize = False
          Enabled = False
          MaxValue = 255
          MinValue = 0
          TabOrder = 2
          Value = 0
          OnChange = edLongueurChange
        end
        object ChampNom: TEdit
          Left = 118
          Top = 7
          Width = 129
          Height = 22
          AutoSize = False
          TabOrder = 1
          OnChange = ChampNomChange
        end
        object ComboBox1: TComboBox
          Left = 294
          Top = 8
          Width = 175
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 4
          OnChange = ComboBox1Change
          Items.Strings = (
            'Alphanum'#233'rique'
            'Num'#233'rique'
            'Date'
            'Heure')
        end
        object GroupBox3: TGroupBox
          Left = 54
          Top = 104
          Width = 345
          Height = 49
          Caption = ' S'#233'lection '
          TabOrder = 5
          object cbTouteValeur: TRadioButton
            Left = 8
            Top = 18
            Width = 89
            Height = 17
            Caption = 'Toute valeur'
            Checked = True
            TabOrder = 0
            TabStop = True
            OnClick = cbTouteValeurClick
          end
          object cbValeurPositive: TRadioButton
            Left = 96
            Top = 18
            Width = 108
            Height = 17
            Caption = 'Valeur positive'
            TabOrder = 1
            OnClick = cbValeurPositiveClick
          end
          object cbValeurNegative: TRadioButton
            Left = 208
            Top = 18
            Width = 108
            Height = 17
            Caption = 'Valeur n'#233'gative'
            TabOrder = 2
            OnClick = cbValeurNegativeClick
          end
        end
        object cbFormatDate: TComboBox
          Left = 294
          Top = 34
          Width = 92
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 6
          Visible = False
          OnChange = cbFormatDateChange
          Items.Strings = (
            'AAAA/MM/JJ'
            'JJ/MM/AAAA'
            'MM/JJ/AAAA'
            'JJMMAAAA'
            'AAAAMMJJ'
            'JJMMAA'
            'JJ/MM/AA')
        end
        object cbType: TComboBox
          Left = 492
          Top = 71
          Width = 92
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 3
          OnChange = cbTypeChange
          Items.Strings = (
            'Normal'
            'Constante'
            'Calcul'
            'R'#233'f'#233'rence'
            'Concat'#233'nation')
        end
        object Nbdecimales: TSpinEdit
          Left = 492
          Top = 101
          Width = 49
          Height = 23
          AutoSize = False
          MaxValue = 0
          MinValue = 0
          TabOrder = 7
          Value = 2
          OnChange = NbdecimalesChange
        end
        object cblibellenul: TComboBox
          Left = 486
          Top = 124
          Width = 74
          Height = 21
          Style = csDropDownList
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clYellow
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 8
          OnChange = cblibellenulChange
          Items.Strings = (
            'OUI m'#234'me si le libelle est nul'
            'NON si le libelle est nul')
        end
        object ComboChampNul: TComboBox
          Left = 452
          Top = 153
          Width = 73
          Height = 21
          Style = csDropDownList
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clYellow
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 9
          Items.Strings = (
            'Defaut'
            'Remplacer les informations existantes'
            'Laisser les informations existantes'
            'Effacer les champs nuls'
            '')
        end
        object Combochampnul1: TComboBox
          Left = 496
          Top = 128
          Width = 74
          Height = 21
          Style = csDropDownList
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clYellow
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 10
          Visible = False
          OnChange = Combochampnul1Change
          Items.Strings = (
            'Remplacer les informations existantes'
            'Laisser les informations existantes')
        end
        object Cacher: TCheckBox
          Left = 480
          Top = 131
          Width = 57
          Height = 17
          Caption = 'Stock'#233
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 11
          Visible = False
        end
        object cbFormatDatesortie: TComboBox
          Left = 492
          Top = 34
          Width = 93
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 12
          Visible = False
          OnChange = cbFormatDatesortieChange
          Items.Strings = (
            'AAAA/MM/JJ'
            'JJ/MM/AAAA'
            'MM/JJ/AAAA'
            'JJMMAAAA'
            'AAAAMMJJ'
            'JJMMAA'
            'JJ/MM/AA')
        end
        object Interne: TCheckBox
          Left = 402
          Top = 130
          Width = 103
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Non Stock'#233
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 13
          OnClick = InterneClick
        end
        object MemoComment: TMemo
          Left = 56
          Top = 155
          Width = 524
          Height = 21
          TabStop = False
          BorderStyle = bsNone
          ReadOnly = True
          TabOrder = 14
          WantReturns = False
        end
        object ARRONDI: TCheckBox
          Left = 538
          Top = 104
          Width = 65
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Arrondir'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 15
          OnClick = ARRONDIClick
        end
      end
      object tsComplement: TTabSheet
        Caption = 'Compl'#233'ment'
        object RadioGroup1: TRadioGroup
          Left = 8
          Top = 2
          Width = 281
          Height = 41
          Caption = ' &Transformation '
          Columns = 3
          ItemIndex = 0
          Items.Strings = (
            'Aucune'
            'Majuscule'
            'Minuscule')
          TabOrder = 0
          OnClick = RadioGroup1Click
        end
        object GroupBox2: TGroupBox
          Left = 8
          Top = 112
          Width = 449
          Height = 65
          TabOrder = 1
          object Label16: TLabel
            Left = 105
            Top = 19
            Width = 69
            Height = 13
            Caption = 'Fichier externe'
          end
          object Label36: TLabel
            Left = 28
            Top = 40
            Width = 127
            Height = 13
            Caption = 'Famille de correspondance'
          end
          object cbTableExterne: TCheckBox
            Left = 9
            Top = 17
            Width = 88
            Height = 17
            Caption = 'Table externe'
            TabOrder = 0
            OnClick = cbTableExterneClick
          end
          object btnEditTbl: TButton
            Left = 358
            Top = 15
            Width = 75
            Height = 20
            Hint = 'Param'#233'trer tables de correspondances'
            Caption = 'Editer'
            TabOrder = 1
            OnClick = btnEditTblClick
          end
          object edNomTable: THCritMaskEdit
            Left = 183
            Top = 15
            Width = 157
            Height = 21
            TabOrder = 2
            OnChange = edNomTableChange
            TagDispatch = 0
            DataType = 'OPENFILE(*.TXT;*.*)'
            ElipsisButton = True
          end
          object cbTableCorr: TCheckBox
            Left = 9
            Top = -1
            Width = 150
            Height = 17
            Caption = 'Table de correspondance'
            TabOrder = 3
            OnClick = cbTableCorrClick
          end
          object Famille: TEdit
            Left = 183
            Top = 38
            Width = 157
            Height = 21
            AutoSize = False
            TabOrder = 4
            OnChange = FamilleChange
          end
        end
        object GroupBox5: TGroupBox
          Left = 8
          Top = 44
          Width = 449
          Height = 61
          Caption = ' Cadrage '
          TabOrder = 2
          object Label8: TLabel
            Left = 262
            Top = 20
            Width = 74
            Height = 13
            Caption = 'Longueur totale'
          end
          object Label18: TLabel
            Left = 168
            Top = 20
            Width = 46
            Height = 13
            Caption = 'Caract'#232're'
          end
          object cbComplement: TCheckBox
            Left = 8
            Top = 19
            Width = 153
            Height = 17
            Caption = 'Compl'#233'ter la zone avec un '
            TabOrder = 0
            OnClick = cbComplementClick
          end
          object edComplLgn: TSpinEdit
            Left = 372
            Top = 16
            Width = 41
            Height = 22
            MaxValue = 0
            MinValue = 0
            TabOrder = 1
            Value = 0
            OnChange = edComplLgnChange
          end
          object edComplCar: TEdit
            Left = 222
            Top = 16
            Width = 18
            Height = 21
            TabOrder = 2
            Text = '0'
            OnChange = edComplCarChange
          end
          object cbAlignLeft: TCheckBox
            Left = 8
            Top = 40
            Width = 121
            Height = 17
            Caption = 'Compl'#233'ter '#224' gauche'
            TabOrder = 3
            OnClick = cbAlignLeftClick
          end
          object CbLngRef: TCheckBox
            Left = 260
            Top = 40
            Width = 125
            Height = 17
            Alignment = taLeftJustify
            Caption = 'Longueur minimun'
            TabOrder = 4
            OnClick = CbLngRefClick
          end
        end
      end
      object tsCALCUL: TTabSheet
        Caption = 'Calcul'
        object Label17: TLabel
          Left = 327
          Top = -1
          Width = 193
          Height = 13
          Caption = 'Liste des champs calcul'#233's et num'#233'riques'
        end
        object SpeedButton23: THSpeedButton
          Left = 157
          Top = 125
          Width = 25
          Height = 25
          Caption = '+'
          OnClick = SpeedButton23Click
        end
        object SpeedButton24: THSpeedButton
          Left = 184
          Top = 125
          Width = 23
          Height = 25
          Caption = '-'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          OnClick = SpeedButton23Click
        end
        object SpeedButton25: THSpeedButton
          Left = 209
          Top = 125
          Width = 24
          Height = 25
          Caption = 'x'
          OnClick = SpeedButton25Click
        end
        object SpeedButton26: THSpeedButton
          Left = 235
          Top = 125
          Width = 25
          Height = 25
          Caption = '/'
          OnClick = SpeedButton23Click
        end
        object SpeedButton27: THSpeedButton
          Left = 262
          Top = 125
          Width = 25
          Height = 25
          Caption = '('
          OnClick = SpeedButton23Click
        end
        object SpeedButton28: THSpeedButton
          Left = 289
          Top = 125
          Width = 25
          Height = 25
          Caption = ')'
          OnClick = SpeedButton23Click
        end
        object Label9: TLabel
          Left = 8
          Top = -1
          Width = 37
          Height = 13
          Caption = 'Formule'
        end
        object SpeedButton29: THSpeedButton
          Left = 317
          Top = 125
          Width = 65
          Height = 25
          Caption = 'ENTRE ( , )'
          OnClick = SpeedButton29Click
        end
        object SpeedButton31: THSpeedButton
          Left = 453
          Top = 125
          Width = 50
          Height = 25
          Caption = 'DATE'
          OnClick = SpeedButton31Click
        end
        object SpeedButton1: THSpeedButton
          Left = 130
          Top = 125
          Width = 25
          Height = 25
          Caption = '='
          OnClick = SpeedButton23Click
        end
        object SpeedButton2: THSpeedButton
          Left = 8
          Top = 125
          Width = 23
          Height = 25
          Caption = 'SI'
          OnClick = SpeedButton23Click
        end
        object SpeedButton14: THSpeedButton
          Left = 32
          Top = 125
          Width = 48
          Height = 25
          Caption = 'ALORS'
          OnClick = SpeedButton23Click
        end
        object SpeedButton15: THSpeedButton
          Left = 81
          Top = 125
          Width = 47
          Height = 25
          Caption = 'SINON'
          OnClick = SpeedButton23Click
        end
        object Label26: TLabel
          Left = 292
          Top = 158
          Width = 255
          Height = 13
          Caption = 'Formule d'#39'exemple : SI ([,]='#39#39') ALORS [,] SINON ([,]+[,])'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CheckCalcul: TToolbarButton97
          Left = 515
          Top = 125
          Width = 27
          Height = 25
          Hint = 'V'#233'rification syntaxe'
          HelpContext = 1710
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          Margin = 5
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          OnClick = CalFormuleExit
          GlobalIndexImage = 'M0013_S16G1'
          IsControl = True
        end
        object SpeedButton37: THSpeedButton
          Left = 57
          Top = 152
          Width = 23
          Height = 25
          Caption = '>'
          OnClick = SpeedButton23Click
        end
        object SpeedButton38: THSpeedButton
          Left = 32
          Top = 152
          Width = 23
          Height = 25
          Caption = '<'
          OnClick = SpeedButton23Click
        end
        object SpeedButton39: THSpeedButton
          Left = 106
          Top = 152
          Width = 22
          Height = 25
          Caption = '>='
          OnClick = SpeedButton23Click
        end
        object SpeedButton40: THSpeedButton
          Left = 82
          Top = 152
          Width = 23
          Height = 25
          Caption = '<='
          OnClick = SpeedButton23Click
        end
        object SpeedButton41: THSpeedButton
          Left = 130
          Top = 152
          Width = 25
          Height = 25
          Caption = 'ET'
          OnClick = SpeedButton23Click
        end
        object SpeedButton42: THSpeedButton
          Left = 157
          Top = 152
          Width = 25
          Height = 25
          Caption = 'OU'
          OnClick = SpeedButton23Click
        end
        object SpeedButton55: THSpeedButton
          Left = 184
          Top = 152
          Width = 76
          Height = 25
          Caption = 'LIGNE(+/-X)'
          OnClick = SpeedButton55Click
        end
        object SpeedButton56: THSpeedButton
          Left = 8
          Top = 152
          Width = 23
          Height = 25
          Caption = '<>'
          OnClick = SpeedButton56Click
        end
        object lbChampNum: TListBox
          Left = 328
          Top = 15
          Width = 223
          Height = 105
          DragMode = dmAutomatic
          ItemHeight = 13
          TabOrder = 0
          OnDblClick = lbChampNumDblClick
          OnEnter = tsCALCULEnter
        end
        object Button2: TButton
          Left = 385
          Top = 125
          Width = 65
          Height = 25
          Caption = 'VAL ( [ , ] ) '
          TabOrder = 1
          OnClick = SpeedButton23Click
        end
        object CalFormule: TColorMemo
          Left = 8
          Top = 15
          Width = 305
          Height = 105
          ScrollBars = ssVertical
          TabOrder = 2
          OnChange = CalFormuleChange
          OnDragDrop = CalFormuleDragDrop
          OnDragOver = CalFormuleDragOver
          OnEnter = CalFormuleEnter
          ColorRules = <>
        end
      end
      object tsREFERENCE: TTabSheet
        Caption = 'R'#233'f'#233'rence'
        object Label10: TLabel
          Left = 8
          Top = 75
          Width = 44
          Height = 13
          Caption = '&Condition'
          FocusControl = RefCondition
        end
        object lbRang: TLabel
          Left = 297
          Top = 149
          Width = 75
          Height = 13
          Caption = 'Ordre de priorit'#233
        end
        object Checkcondition: TToolbarButton97
          Left = 509
          Top = 70
          Width = 29
          Height = 22
          Hint = 'V'#233'rification syntaxe'
          HelpContext = 1710
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          Margin = 5
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          OnClick = RefConditionExit
          GlobalIndexImage = 'M0013_S16G1'
          IsControl = True
        end
        object SpeedButton43: THSpeedButton
          Left = 8
          Top = 97
          Width = 25
          Height = 25
          Caption = '<'
          OnClick = SpeedButton43Click
        end
        object SpeedButton44: THSpeedButton
          Left = 38
          Top = 97
          Width = 25
          Height = 25
          Caption = '>'
          OnClick = SpeedButton43Click
        end
        object SpeedButton45: THSpeedButton
          Left = 68
          Top = 97
          Width = 25
          Height = 25
          Caption = '='
          OnClick = SpeedButton43Click
        end
        object SpeedButton46: THSpeedButton
          Left = 396
          Top = 97
          Width = 85
          Height = 25
          Caption = 'VAL ( [ , ] )'
          OnClick = SpeedButton43Click
        end
        object SpeedButton47: THSpeedButton
          Left = 99
          Top = 97
          Width = 25
          Height = 25
          Caption = '<>'
          OnClick = SpeedButton43Click
        end
        object SpeedButton48: THSpeedButton
          Left = 128
          Top = 97
          Width = 25
          Height = 25
          Caption = '<='
          OnClick = SpeedButton43Click
        end
        object SpeedButton49: THSpeedButton
          Left = 157
          Top = 97
          Width = 25
          Height = 25
          Caption = '>='
          OnClick = SpeedButton43Click
        end
        object SpeedButton50: THSpeedButton
          Left = 187
          Top = 97
          Width = 25
          Height = 25
          Caption = '('
          OnClick = SpeedButton43Click
        end
        object SpeedButton51: THSpeedButton
          Left = 216
          Top = 97
          Width = 25
          Height = 25
          Caption = ')'
          OnClick = SpeedButton43Click
        end
        object SpeedButton52: THSpeedButton
          Left = 247
          Top = 97
          Width = 25
          Height = 25
          Caption = 'ET'
          OnClick = SpeedButton43Click
        end
        object SpeedButton53: THSpeedButton
          Left = 276
          Top = 97
          Width = 25
          Height = 25
          Caption = 'OU'
          OnClick = SpeedButton43Click
        end
        object SpeedButton54: THSpeedButton
          Left = 306
          Top = 97
          Width = 85
          Height = 25
          Caption = 'ENTRE ( , )'
          OnClick = SpeedButton43Click
        end
        object RefPosterieur: TCheckBox
          Left = 397
          Top = 147
          Width = 124
          Height = 17
          Caption = '&R'#233'f'#233'rence post'#233'rieure'
          TabOrder = 0
          OnClick = RefPosterieurClick
        end
        object RefCondition: TEdit
          Left = 60
          Top = 70
          Width = 441
          Height = 22
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Courier New'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnChange = RefConditionChange
          OnEnter = RefConditionEnter
        end
        object Memo5: TMemo
          Left = 8
          Top = 8
          Width = 529
          Height = 58
          BorderStyle = bsNone
          Enabled = False
          Lines.Strings = (
            
              'Une r'#233'f'#233'rence permet de r'#233'cup'#233'rer une valeur sur une ligne non s' +
              #233'lectionn'#233'e par la condition compl'#233'mentaire.'
            
              'Exemple: la partie de l'#39'enregistrement concern'#233' se trouve en pos' +
              'ition 30 apr'#232's le mot '#39'COMPTE'#39' '
            'qui se trouve lui-m'#234'me en position 21 sur 7 caract'#232'res.'
            
              'on '#233'crira dans la condition : [ 21,6 ] = '#39'COMPTE'#39' puis on affect' +
              'era le champ ainsi : [21,7]')
          TabOrder = 2
        end
        object GroupBox4: TGroupBox
          Left = 8
          Top = 126
          Width = 223
          Height = 49
          Caption = ' &Zone '#224' r'#233'cup'#233'rer '
          TabOrder = 3
          object Label11: TLabel
            Left = 8
            Top = 23
            Width = 37
            Height = 13
            Caption = '&Position'
            FocusControl = RefPosition
          end
          object Label12: TLabel
            Left = 112
            Top = 23
            Width = 48
            Height = 13
            Caption = '&Longueur '
            FocusControl = RefLongueur
          end
          object RefPosition: TEdit
            Left = 50
            Top = 19
            Width = 41
            Height = 21
            TabOrder = 0
            OnChange = RefPositionChange
            OnEnter = RefPositionEnter
          end
          object RefLongueur: TEdit
            Left = 168
            Top = 19
            Width = 41
            Height = 21
            TabOrder = 1
            OnChange = RefLongueurChange
            OnEnter = RefLongueurEnter
          end
        end
        object EdRang: TSpinEdit
          Left = 252
          Top = 145
          Width = 41
          Height = 22
          MaxValue = 0
          MinValue = 0
          TabOrder = 4
          Value = 0
          OnChange = EdRangChange
        end
      end
      object tsCONSTANTE: TTabSheet
        Caption = 'Constante'
        object Valeur: TLabel
          Left = 8
          Top = 76
          Width = 30
          Height = 13
          Caption = 'Valeur'
        end
        object Pas: TLabel
          Left = 8
          Top = 101
          Width = 18
          Height = 13
          Caption = 'Pas'
        end
        object Longueur: TLabel
          Left = 8
          Top = 125
          Width = 45
          Height = 13
          Caption = 'Longueur'
        end
        object CstVal: TEdit
          Left = 64
          Top = 72
          Width = 121
          Height = 21
          TabOrder = 0
          Text = 'CstValeur'
          OnChange = CstValChange
        end
        object CstPas: TSpinEdit
          Left = 64
          Top = 96
          Width = 49
          Height = 22
          MaxValue = 99999999
          MinValue = 0
          TabOrder = 1
          Value = 0
          OnChange = CstPasChange
        end
        object CstLon: TSpinEdit
          Left = 64
          Top = 120
          Width = 49
          Height = 22
          MaxValue = 255
          MinValue = 0
          TabOrder = 2
          Value = 0
          OnChange = CstLonChange
        end
        object Memo4: TMemo
          Left = 8
          Top = 8
          Width = 425
          Height = 57
          TabStop = False
          BorderStyle = bsNone
          ReadOnly = True
          TabOrder = 3
          WantReturns = False
        end
      end
      object TabConditioncomp: TTabSheet
        Caption = 'Condition compl'#233'mentaire'
        DesignSize = (
          680
          178)
        object SpeedButton16: THSpeedButton
          Left = 12
          Top = 94
          Width = 25
          Height = 25
          Caption = '<'
          OnClick = SpeedButton16Click
        end
        object SpeedButton17: THSpeedButton
          Left = 42
          Top = 94
          Width = 25
          Height = 25
          Caption = '>'
          OnClick = SpeedButton16Click
        end
        object SpeedButton18: THSpeedButton
          Left = 71
          Top = 94
          Width = 25
          Height = 25
          Caption = '='
          OnClick = SpeedButton16Click
        end
        object SpeedButton19: THSpeedButton
          Left = 100
          Top = 94
          Width = 25
          Height = 25
          Caption = '<>'
          OnClick = SpeedButton16Click
        end
        object SpeedButton20: THSpeedButton
          Left = 129
          Top = 94
          Width = 25
          Height = 25
          Caption = '<='
          OnClick = SpeedButton16Click
        end
        object SpeedButton21: THSpeedButton
          Left = 158
          Top = 94
          Width = 25
          Height = 25
          Caption = '>='
          OnClick = SpeedButton16Click
        end
        object SpeedButton22: THSpeedButton
          Left = 247
          Top = 94
          Width = 25
          Height = 25
          Caption = 'ET'
          OnClick = SpeedButton16Click
        end
        object SpeedButton32: THSpeedButton
          Left = 276
          Top = 94
          Width = 25
          Height = 25
          Caption = 'OU'
          OnClick = SpeedButton16Click
        end
        object SpeedButton33: THSpeedButton
          Left = 305
          Top = 94
          Width = 85
          Height = 25
          Caption = 'ENTRE ( , )'
          OnClick = SpeedButton16Click
        end
        object SpeedButton34: THSpeedButton
          Left = 188
          Top = 94
          Width = 25
          Height = 25
          Caption = '('
          OnClick = SpeedButton16Click
        end
        object SpeedButton35: THSpeedButton
          Left = 217
          Top = 94
          Width = 25
          Height = 25
          Caption = ')'
          OnClick = SpeedButton16Click
        end
        object SpeedButton36: THSpeedButton
          Left = 394
          Top = 94
          Width = 85
          Height = 25
          Caption = 'VAL ( [ , ] )'
          OnClick = SpeedButton16Click
        end
        object Checksyntaxe: TToolbarButton97
          Left = 393
          Top = 125
          Width = 27
          Height = 22
          Hint = 'V'#233'rification syntaxe'
          HelpContext = 1710
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          Margin = 5
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          OnClick = ChecksyntaxeClick
          GlobalIndexImage = 'M0013_S16G1'
          IsControl = True
        end
        object Label30: TLabel
          Left = 8
          Top = 159
          Width = 52
          Height = 13
          Caption = 'Ordre de tri'
        end
        object ApercuConditionComp: TToolbarButton97
          Left = 662
          Top = 93
          Width = 25
          Height = 25
          Hint = 'Aper'#231'u condition'
          AllowAllUp = True
          Anchors = [akTop, akRight]
          ParentShowHint = False
          ShowHint = True
          OnClick = ApercuConditionCompClick
          GlobalIndexImage = 'Z0138_S16G1'
        end
        object Memo3: THRichEditOLE
          Left = 0
          Top = 0
          Width = 680
          Height = 90
          Align = alTop
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnChange = Memo3Change
          OnEnter = Memo3Enter
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
              'l;}}'
            
              '{\*\generator Msftedit 5.41.21.2507;}\viewkind4\uc1\pard\f0\fs16' +
              ' '
            '\par }')
        end
        object bCondition: TCheckBox
          Left = 8
          Top = 128
          Width = 236
          Height = 17
          Caption = 'Ne pas tenir compte de la condition g'#233'n'#233'rale'
          TabOrder = 1
          OnClick = bConditionClick
        end
        object EnregUnique: TCheckBox
          Left = 252
          Top = 128
          Width = 137
          Height = 17
          Caption = 'Enregistrement unique'
          TabOrder = 2
          OnClick = EnregUniqueClick
        end
        object ORDER: TEdit
          Left = 69
          Top = 155
          Width = 459
          Height = 21
          TabOrder = 3
          OnChange = ORDERChange
        end
      end
      object tsCoherence: TTabSheet
        Caption = 'Coh'#233'rence'
        object MemoMsg: TMemo
          Left = 8
          Top = 184
          Width = 329
          Height = 49
          TabOrder = 0
          Visible = False
        end
        object tvchampinvisible: TTreeView
          Left = 0
          Top = 0
          Width = 680
          Height = 178
          Align = alClient
          DragMode = dmAutomatic
          HideSelection = False
          Images = ImageList1
          Indent = 19
          ParentShowHint = False
          ShowHint = False
          TabOrder = 1
          OnDeletion = tvchampinvisibleDeletion
        end
      end
      object Tabparam: TTabSheet
        Caption = 'Entr'#233'e'
        ImageIndex = 7
        object LFichier: TLabel
          Left = 14
          Top = 105
          Width = 80
          Height = 13
          Caption = 'Fichier '#224' importer'
        end
        object Label28: TLabel
          Tag = 100
          Left = 385
          Top = 144
          Width = 49
          Height = 13
          Caption = 'D'#233'limiteur '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          Visible = False
        end
        object Label24: TLabel
          Left = 385
          Top = 105
          Width = 51
          Height = 13
          Caption = 'Caract'#232'res'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object FILENAME: THCritMaskEdit
          Left = 100
          Top = 101
          Width = 264
          Height = 21
          TabOrder = 1
          OnChange = FILENAMEChange
          TagDispatch = 0
          DataType = 'OPENFILE(*.TXT;*.ZIP;*.TRT;*.*)'
          ElipsisButton = True
        end
        object GroupBox6: TGroupBox
          Left = 8
          Top = 0
          Width = 553
          Height = 90
          TabOrder = 0
          object Label25: TLabel
            Left = 6
            Top = 16
            Width = 42
            Height = 13
            Caption = 'Domaine'
          end
          object Label19: TLabel
            Left = 178
            Top = 16
            Width = 25
            Height = 13
            Caption = 'Code'
          end
          object Label22: TLabel
            Left = 6
            Top = 40
            Width = 30
            Height = 13
            Caption = 'Libell'#233
          end
          object Label39: TLabel
            Left = 6
            Top = 60
            Width = 58
            Height = 13
            Caption = 'Compl'#233'ment'
            FocusControl = EdNature
          end
          object Chargement: TToolbarButton97
            Left = 327
            Top = 12
            Width = 28
            Height = 21
            Hint = 'Chargement'
            HelpContext = 1710
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Margin = 3
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = ChargementClick
            GlobalIndexImage = 'Z0205_S16G1'
            IsControl = True
          end
          object Label37: TLabel
            Left = 360
            Top = 40
            Width = 32
            Height = 13
            Caption = 'Nature'
          end
          object Label38: TLabel
            Left = 360
            Top = 64
            Width = 33
            Height = 13
            Caption = 'Editeur'
          end
          object TPAYS: TLabel
            Left = 360
            Top = 16
            Width = 23
            Height = 13
            Caption = 'Pays'
          end
          object Fichesauve: THCritMaskEdit
            Left = 212
            Top = 12
            Width = 112
            Height = 21
            MaxLength = 12
            TabOrder = 0
            OnChange = FichesauveChange
            OnExit = FichesauveExit
            TagDispatch = 0
            OnElipsisClick = FichesauveElipsisClick
          end
          object edcomment: TEdit
            Left = 69
            Top = 36
            Width = 283
            Height = 21
            MaxLength = 60
            TabOrder = 2
          end
          object EdNature: THCritMaskEdit
            Left = 69
            Top = 60
            Width = 284
            Height = 21
            MaxLength = 12
            TabOrder = 4
            OnChange = FichesauveChange
            OnExit = FichesauveExit
            TagDispatch = 0
            OnElipsisClick = FichesauveElipsisClick
          end
          object Typetransfert: THCritMaskEdit
            Left = 426
            Top = 37
            Width = 112
            Height = 21
            MaxLength = 12
            TabOrder = 1
            OnChange = FichesauveChange
            OnExit = FichesauveExit
            TagDispatch = 0
            OnElipsisClick = FichesauveElipsisClick
          end
          object Editeur: THCritMaskEdit
            Left = 426
            Top = 61
            Width = 112
            Height = 21
            MaxLength = 12
            TabOrder = 3
            OnChange = FichesauveChange
            OnExit = FichesauveExit
            TagDispatch = 0
            OnElipsisClick = FichesauveElipsisClick
          end
          object Domaine: THValComboBox
            Left = 69
            Top = 13
            Width = 104
            Height = 21
            Style = csDropDownList
            ItemHeight = 0
            TabOrder = 5
            TagDispatch = 0
            ComboWidth = 300
          end
          object BPAYS: THValComboBox
            Left = 426
            Top = 11
            Width = 111
            Height = 21
            Style = csDropDownList
            BiDiMode = bdLeftToRight
            ItemHeight = 0
            ParentBiDiMode = False
            TabOrder = 6
            TagDispatch = 0
            Vide = True
            DataType = 'TTPAYS'
            ComboWidth = 150
          end
        end
        object RTypefichier: TRadioGroup
          Left = 175
          Top = 132
          Width = 185
          Height = 37
          Caption = 'Type de fichier'
          Columns = 2
          ItemIndex = 0
          Items.Strings = (
            'ASCII Fixe'
            'ASCII D'#233'limit'#233)
          TabOrder = 4
          Visible = False
          OnClick = RTypefichierClick
        end
        object cbDelimChamp: TComboBox
          Tag = 100
          Left = 444
          Top = 140
          Width = 81
          Height = 21
          ItemHeight = 13
          TabOrder = 5
          Visible = False
          Items.Strings = (
            'point virgule ( ; )'
            'virgule ( , )'
            'Tabulation ( TAB )'
            'Barre vertical ( | )'
            'Deux points ( : )'
            'Autres')
        end
        object ModeParam: TRadioGroup
          Left = 7
          Top = 132
          Width = 157
          Height = 37
          Caption = 'Mode de Param'#233'trage'
          Columns = 2
          ItemIndex = 0
          Items.Strings = (
            'Complet'
            'Simplifi'#233)
          TabOrder = 3
          OnClick = ModeParamClick
        end
        object cbTypeCar: TComboBox
          Left = 444
          Top = 101
          Width = 79
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 2
          OnChange = cbTypeCarChange
          Items.Strings = (
            'OEM'
            'ANSI')
        end
      end
      object tsCHPCONDITION: TTabSheet
        Caption = 'Condition g'#233'n'#233'rale'
        ImageIndex = 8
        object SpeedButton3: THSpeedButton
          Left = 9
          Top = 99
          Width = 25
          Height = 25
          Caption = '<'
          OnClick = SpeedButton3Click
        end
        object SpeedButton4: THSpeedButton
          Left = 38
          Top = 99
          Width = 25
          Height = 25
          Caption = '>'
          OnClick = SpeedButton3Click
        end
        object SpeedButton5: THSpeedButton
          Left = 67
          Top = 99
          Width = 25
          Height = 25
          Caption = '='
          OnClick = SpeedButton3Click
        end
        object SpeedButton6: THSpeedButton
          Left = 96
          Top = 99
          Width = 25
          Height = 25
          Caption = '<>'
          OnClick = SpeedButton3Click
        end
        object SpeedButton8: THSpeedButton
          Left = 125
          Top = 99
          Width = 25
          Height = 25
          Caption = '<='
          OnClick = SpeedButton3Click
        end
        object SpeedButton9: THSpeedButton
          Left = 154
          Top = 99
          Width = 25
          Height = 25
          Caption = '>='
          OnClick = SpeedButton3Click
        end
        object SpeedButton10: THSpeedButton
          Left = 241
          Top = 99
          Width = 25
          Height = 25
          Caption = 'ET'
          OnClick = SpeedButton3Click
        end
        object SpeedButton11: THSpeedButton
          Left = 270
          Top = 99
          Width = 25
          Height = 25
          Caption = 'OU'
          OnClick = SpeedButton3Click
        end
        object SpeedButton7: THSpeedButton
          Left = 299
          Top = 99
          Width = 85
          Height = 25
          Caption = 'ENTRE ( , )'
          OnClick = SpeedButton3Click
        end
        object SpeedButton12: THSpeedButton
          Left = 183
          Top = 99
          Width = 25
          Height = 25
          Caption = '('
          OnClick = SpeedButton3Click
        end
        object SpeedButton13: THSpeedButton
          Left = 212
          Top = 99
          Width = 25
          Height = 25
          Caption = ')'
          OnClick = SpeedButton3Click
        end
        object SpeedButton30: THSpeedButton
          Left = 388
          Top = 99
          Width = 85
          Height = 25
          Caption = 'VAL ( [ , ] )'
          OnClick = SpeedButton3Click
        end
        object memCondition: THRichEditOLE
          Left = 0
          Top = 0
          Width = 680
          Height = 90
          Align = alTop
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnEnter = memConditionEnter
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
              'l;}}'
            
              '{\*\generator Msftedit 5.41.21.2507;}\viewkind4\uc1\pard\f0\fs16' +
              ' memCondition'
            '\par }')
        end
      end
      object TabSortie: TTabSheet
        Caption = 'Commande/Sortie'
        ImageIndex = 9
        object Label31: TLabel
          Left = 17
          Top = 114
          Width = 76
          Height = 13
          Caption = 'Fichier de Sortie'
        end
        object Label32: TLabel
          Left = 17
          Top = 144
          Width = 84
          Height = 13
          Caption = 'Requ'#234'te suivante'
        end
        object Radiotable: TRadioGroup
          Left = 8
          Top = 53
          Width = 225
          Height = 41
          Hint = 'Type de donn'#233'es'
          Caption = 'Type de table'
          Columns = 2
          ItemIndex = 0
          Items.Strings = (
            'Standard'
            'ASCII')
          TabOrder = 1
          OnClick = RadiotableClick
        end
        object FicheSortie: THCritMaskEdit
          Left = 139
          Top = 110
          Width = 243
          Height = 21
          TabOrder = 2
          OnChange = FicheSortieChange
          TagDispatch = 0
          DataType = 'SAVEFILE(*.TXT;*.TRA;*.TRT;*.*)'
          ElipsisButton = True
        end
        object Table: THCritMaskEdit
          Left = 139
          Top = 140
          Width = 156
          Height = 21
          MaxLength = 12
          TabOrder = 3
          OnChange = TableChange
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = TableElipsisClick
        end
        object GroupBox11: TGroupBox
          Left = 8
          Top = 1
          Width = 505
          Height = 49
          Caption = 'Pr'#233'-traitement'
          TabOrder = 0
          DesignSize = (
            505
            49)
          object Label35: TLabel
            Left = 12
            Top = 21
            Width = 118
            Height = 13
            Caption = 'Executer une commande'
          end
          object ToolbarButton971: TToolbarButton97
            Left = 464
            Top = 17
            Width = 26
            Height = 21
            Hint = 'Valider'
            Alignment = taLeftJustify
            Anchors = [akRight]
            DisplayMode = dmGlyphOnly
            Caption = 'Executer'
            Flat = False
            NoBorder = True
            OnClick = ToolbarButton971Click
            GlobalIndexImage = 'Z0217_S16G1'
          end
          object EdShellExecute: THCritMaskEdit
            Left = 140
            Top = 17
            Width = 316
            Height = 21
            TabOrder = 0
            OnChange = EdShellExecuteChange
            TagDispatch = 0
            DataType = 'OPENFILE(*.EXE;*.*)'
            ElipsisButton = True
          end
        end
      end
      object TabCorresp: TTabSheet
        Caption = 'Table de correspondance'
        ImageIndex = 11
        object ListeProfile: TListBox
          Left = 50
          Top = 48
          Width = 312
          Height = 122
          DragMode = dmAutomatic
          ItemHeight = 13
          TabOrder = 0
        end
        object GroupBox7: TGroupBox
          Left = 24
          Top = 2
          Width = 353
          Height = 39
          TabOrder = 1
          DesignSize = (
            353
            39)
          object LTable: TLabel
            Left = 6
            Top = 14
            Width = 28
            Height = 13
            Caption = 'Profils'
          end
          object bDefaire: TToolbarButton97
            Left = 301
            Top = 9
            Width = 28
            Height = 23
            Hint = 'Vider la liste'
            Alignment = taLeftJustify
            AllowAllUp = True
            DisplayMode = dmGlyphOnly
            Caption = 'Vider la liste'
            Flat = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            ParentFont = False
            Spacing = -1
            OnClick = bDefaireClick
            GlobalIndexImage = 'Z0080_S16G1'
            IsControl = True
          end
          object BValide: TToolbarButton97
            Left = 261
            Top = 10
            Width = 29
            Height = 21
            Hint = 'Ouvrir'
            Alignment = taLeftJustify
            Anchors = [akRight]
            DisplayMode = dmGlyphOnly
            Caption = 'Ouvrir'
            Flat = False
            NoBorder = True
            OnClick = BValideClick
            GlobalIndexImage = 'Z0217_S16G1'
          end
          object Profile: THCritMaskEdit
            Left = 52
            Top = 10
            Width = 197
            Height = 21
            MaxLength = 12
            TabOrder = 0
            TagDispatch = 0
            ElipsisButton = True
            OnElipsisClick = ProfileElipsisClick
          end
        end
      end
      object TabLien: TTabSheet
        Caption = 'Liens inter section'
        ImageIndex = 11
        OnEnter = TabLienEnter
        object Label33: TLabel
          Left = 6
          Top = 7
          Width = 27
          Height = 13
          Caption = 'Table'
        end
        object Label34: TLabel
          Left = 187
          Top = 7
          Width = 72
          Height = 13
          Caption = 'Nom du champ'
        end
        object BValideLien: TToolbarButton97
          Left = 471
          Top = 5
          Width = 29
          Height = 20
          Hint = 'Ouvrir'
          Alignment = taLeftJustify
          DisplayMode = dmGlyphOnly
          Caption = 'Ouvrir'
          Flat = False
          NoBorder = True
          OnClick = BValideLienClick
          GlobalIndexImage = 'Z0217_S16G1'
        end
        object BEffLien: TToolbarButton97
          Left = 500
          Top = 99
          Width = 28
          Height = 22
          Hint = 'Vider la liste'
          Alignment = taLeftJustify
          AllowAllUp = True
          DisplayMode = dmGlyphOnly
          Caption = 'Vider la liste'
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Spacing = -1
          OnClick = BEffLienClick
          GlobalIndexImage = 'Z0080_S16G1'
          IsControl = True
        end
        object LIENINTER: THRichEditOLE
          Left = 0
          Top = 32
          Width = 530
          Height = 55
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnChange = LIENINTERChange
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
              'l;}}'
            
              '{\*\generator Msftedit 5.41.21.2507;}\viewkind4\uc1\pard\f0\fs16' +
              ' '
            '\par }')
        end
        object Memo2: TMemo
          Left = 3
          Top = 99
          Width = 488
          Height = 69
          BorderStyle = bsNone
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowFrame
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Lines.Strings = (
            
              'Indiquer la table section + '#39'/'#39' + les champs s'#233'par'#233's par une vir' +
              'gule '
            'Exemple : Analytique/codejournal,datemouvement,comptegeneral'
            'Le nom des champs doit '#234'tre identique entre les deux '
            'enregistrements.')
          ParentFont = False
          TabOrder = 1
        end
        object TableN: THValComboBox
          Left = 40
          Top = 4
          Width = 140
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 2
          OnExit = TableNExit
          Items.Strings = (
            'essai'
            'essai2')
          TagDispatch = 0
          DisableTab = True
        end
        object CodeChamp: THValComboBox
          Left = 268
          Top = 4
          Width = 182
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 3
          Items.Strings = (
            'essai'
            'essai2')
          TagDispatch = 0
          DisableTab = True
        end
      end
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 0
    Width = 149
    Height = 410
    Align = alLeft
    Caption = 'Panel4'
    TabOrder = 2
    object tvChamp: TTreeView
      Left = 1
      Top = 1
      Width = 147
      Height = 408
      Align = alClient
      DragMode = dmAutomatic
      HideSelection = False
      Images = ImageList1
      Indent = 19
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      OnChange = tvChampChange
      OnChanging = tvChampChanging
      OnClick = tvChampClick
      OnDblClick = tvChampDblClick
      OnDeletion = tvChampDeletion
      OnEditing = tvChampEditing
      OnEnter = tvChampEnter
      OnKeyDown = tvChampKeyDown
      OnKeyPress = tvChampKeyPress
      OnMouseDown = tvChampMouseDown
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Attention : validation non effectu'#233'e !'
      
        '1;?caption?;Aucune '#233'ch'#233'ance n'#39'est s'#233'lectionn'#233'e, op'#233'ration imposs' +
        'ible.;W;O;O;O;'
      
        '2;?caption?;Attention : certaines '#233'ch'#233'ances n'#39'ont pas '#233't'#233' affect' +
        #233'es, confirmer la validation ?;Q;YN;Y;N;'
      '')
    Left = 75
    Top = 119
  end
  object ImageList1: THImageList
    GlobalIndexImages.Strings = (
      'Z2396_S16G1'
      'Z0007_S16G1'
      'Z0007_S16G1'
      'Z0437_S16G1'
      'Z0033_S16G1'
      'Z0034_S16G1'
      'Z0034_S16G1'
      'Z2395_S16G1'
      'Z0007_S16G1'
      'Z0241_S16G1')
    Left = 24
    Top = 132
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 100
    Top = 212
    object Nouveau1: TMenuItem
      Caption = '&Nouveau'
      object Champnormal1: TMenuItem
        Caption = 'Champ fin de liste'
        OnClick = Champnormal1Click
      end
      object Champ1: TMenuItem
        Caption = 'Insertion Champ '
        OnClick = miNouveauClick
      end
      object Concat1: TMenuItem
        Caption = 'Nouvelle partie'
        OnClick = Nouvellepartie1Click
      end
    end
    object Type1: TMenuItem
      Caption = '&Type'
      object miNormal: TMenuItem
        Caption = '&Normal'
        OnClick = miNormalClick
      end
      object miConstante: TMenuItem
        Caption = '&Constante'
        OnClick = miConstanteClick
      end
      object miCalcul: TMenuItem
        Caption = 'C&alcul'
        OnClick = miCalculClick
      end
      object miReference: TMenuItem
        Caption = '&R'#233'f'#233'rence'
        OnClick = miReferenceClick
      end
    end
    object Scruter1: TMenuItem
      Caption = 'Scruter'
      Hint = 'Visualisation du r'#233'sultat'
      OnClick = Scruter1Click
    end
    object N6: TMenuItem
      Caption = '-'
    end
    object Supprimer1: TMenuItem
      Caption = 'Supprimer'
      OnClick = Supprimer1Click
    end
    object Copier1: TMenuItem
      Caption = 'Copier'
      OnClick = Copier1Click
    end
    object Couper1: TMenuItem
      Caption = 'Couper'
      OnClick = Couper1Click
    end
    object Coller1: TMenuItem
      Caption = 'Coller'
      OnClick = Coller1Click
    end
    object N1: TMenuItem
      Caption = '-'
      Visible = False
    end
    object ViewFScript1: TMenuItem
      Caption = 'View FScript'
      OnClick = ViewFScript1Click
    end
    object Imprimer1: TMenuItem
      Caption = 'Imprimer'
      Hint = 'Impression des param'#232'tres'
      OnClick = Imprimer1Click
    end
  end
  object PopupMenu2: TPopupMenu
    Left = 24
    Top = 77
    object Affecter1: TMenuItem
      Caption = 'Affecter Champ'
      OnClick = btnAffecterClick
    end
    object Rejet: TMenuItem
      Caption = 'Rejet ligne'
      OnClick = RejetClick
    end
    object Affecterligne1: TMenuItem
      Caption = 'Affecter ligne'
      OnClick = Affecterligne1Click
    end
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 48
    Top = 248
  end
end

object FParametre: TFParametre
  Left = 309
  Top = 169
  Width = 654
  Height = 480
  Caption = 'Gestion des param'#232'tres'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 205
    Top = 0
    Height = 409
  end
  object TV: TTreeView
    Left = 0
    Top = 0
    Width = 205
    Height = 409
    Align = alLeft
    DragMode = dmAutomatic
    HideSelection = False
    Images = ImageList1
    Indent = 19
    ReadOnly = True
    StateImages = ImageList1
    TabOrder = 0
    OnChange = TVChange
    OnClick = TVClick
  end
  object Panel1: TPanel
    Left = 0
    Top = 409
    Width = 646
    Height = 37
    Align = alBottom
    TabOrder = 1
    DesignSize = (
      646
      37)
    object Binsert: TToolbarButton97
      Left = 261
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Inserer au niveau inf'#233'rieur'
      DisplayMode = dmGlyphOnly
      Caption = 'Nouveau'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BinsertClick
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object BDelete: TToolbarButton97
      Left = 289
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Supprimer'
      DisplayMode = dmGlyphOnly
      Caption = 'Supprimer'
      Flat = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BDeleteClick
      GlobalIndexImage = 'Z0005_S16G1'
    end
    object BAnnuler: TToolbarButton97
      Left = 550
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Anchors = [akTop, akRight]
      Cancel = True
      DisplayMode = dmGlyphOnly
      Caption = 'Fermer'
      Flat = False
      Layout = blGlyphTop
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = BAnnulerClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: TToolbarButton97
      Left = 610
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Aide'
      Anchors = [akTop, akRight]
      DisplayMode = dmGlyphOnly
      Caption = 'Aide'
      Flat = False
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      Visible = False
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object BOuvrir: TToolbarButton97
      Left = 580
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Enregistrer'
      Anchors = [akTop, akRight]
      DisplayMode = dmGlyphOnly
      Caption = 'Ouvrir'
      Flat = False
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      OnClick = BOuvrirClick
      GlobalIndexImage = 'Z0184_S16G1'
    end
    object BRechercher: TToolbarButton97
      Left = 11
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Rechercher dans la liste'
      DisplayMode = dmGlyphOnly
      Caption = 'Chercher'
      Flat = False
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      OnClick = BRechercherClick
      GlobalIndexImage = 'Z0077_S16G1'
    end
    object BImprimer: TToolbarButton97
      Left = 39
      Top = 5
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
      Layout = blGlyphTop
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BImprimerClick
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BFirst: TToolbarButton97
      Left = 88
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Premier'
      Flat = False
      Visible = False
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrev: TToolbarButton97
      Left = 116
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Pr'#233'c'#233'dent'
      Flat = False
      Visible = False
      OnClick = BPrevClick
      GlobalIndexImage = 'Z0017_S16G1'
    end
    object BNext: TToolbarButton97
      Left = 144
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Suivant'
      Flat = False
      Visible = False
      OnClick = BNextClick
      GlobalIndexImage = 'Z0031_S16G1'
    end
    object BLast: TToolbarButton97
      Left = 172
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Dernier'
      Flat = False
      Visible = False
      OnClick = BLastClick
      GlobalIndexImage = 'Z0264_S16G1'
    end
    object bDefaire: TToolbarButton97
      Left = 205
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Annuler les modifications'
      AllowAllUp = True
      DisplayMode = dmGlyphOnly
      Caption = 'Annuler'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BDefaireClick
      GlobalIndexImage = 'Z0075_S16G1'
      IsControl = True
    end
    object bPost: TToolbarButton97
      Left = 233
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Valider les modifications'
      AllowAllUp = True
      Flat = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BPostClick
      GlobalIndexImage = 'Z0127_S16G1'
    end
    object bDupliquer: TToolbarButton97
      Left = 317
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Dupliquer'
      AllowAllUp = True
      Flat = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bDupliquerClick
      GlobalIndexImage = 'Z0038_S16G1'
    end
  end
  object Panel2: TPanel
    Left = 208
    Top = 0
    Width = 438
    Height = 409
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object Splitter2: TSplitter
      Left = 0
      Top = 182
      Width = 438
      Height = 3
      Cursor = crVSplit
      Align = alBottom
    end
    object Panel_Champs: TPanel
      Left = 0
      Top = 185
      Width = 438
      Height = 224
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 0
      object PCPied: TPageControl
        Left = 0
        Top = 0
        Width = 438
        Height = 224
        ActivePage = Page1
        Align = alClient
        TabHeight = 17
        TabOrder = 0
        TabStop = False
        TabWidth = 80
        object Page1: TTabSheet
          Caption = 'Donn'#233'es'
          object TMEContenu: TLabel
            Left = 192
            Top = 77
            Width = 30
            Height = 13
            Caption = 'Valeur'
            FocusControl = MEContenu
          end
          object TMELongueur: TLabel
            Left = 6
            Top = 53
            Width = 45
            Height = 13
            Caption = 'Longueur'
          end
          object TMEMajusMin: TLabel
            Left = 6
            Top = 77
            Width = 45
            Height = 13
            Caption = 'Maj./Min.'
          end
          object TMEType: TLabel
            Left = 192
            Top = 53
            Width = 24
            Height = 13
            Caption = 'Type'
          end
          object TMEContenuoblig: TLabel
            Left = 192
            Top = 101
            Width = 81
            Height = 13
            Caption = 'Valeurs possibles'
            FocusControl = MEContenuoblig
          end
          object TMENomPGI: TLabel
            Left = 192
            Top = 125
            Width = 43
            Height = 13
            Caption = 'Nom PGI'
            FocusControl = MENomPGI
          end
          object TMECommentaire: TLabel
            Left = 6
            Top = 173
            Width = 61
            Height = 13
            Caption = 'Commentaire'
            FocusControl = MECommentaire
          end
          object TMEDomaine: TLabel
            Left = 8
            Top = 5
            Width = 42
            Height = 13
            Caption = 'Domaine'
          end
          object TMETable: TLabel
            Left = 6
            Top = 29
            Width = 27
            Height = 13
            Caption = 'Table'
            FocusControl = METableName
          end
          object TMENom: TLabel
            Left = 192
            Top = 29
            Width = 72
            Height = 13
            Caption = 'Nom du champ'
            FocusControl = MENom
          end
          object Label10: TLabel
            Left = 6
            Top = 124
            Width = 52
            Height = 13
            Caption = 'Alignement'
            FocusControl = MEAlignement
          end
          object Label11: TLabel
            Left = 6
            Top = 148
            Width = 52
            Height = 13
            Caption = 'S'#233'parateur'
            FocusControl = MESeparateur
          end
          object Label12: TLabel
            Left = 192
            Top = 149
            Width = 79
            Height = 13
            Caption = 'Nb de d'#233'cimales'
          end
          object LTEOffset: TLabel
            Left = 192
            Top = 5
            Width = 70
            Height = 13
            Caption = 'Offset (Base 1)'
            FocusControl = TEOffset
          end
          object MEContenu: TMaskEdit
            Left = 278
            Top = 73
            Width = 121
            Height = 21
            MaxLength = 50
            TabOrder = 8
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MEContenuoblig: TMaskEdit
            Left = 278
            Top = 97
            Width = 121
            Height = 21
            MaxLength = 50
            TabOrder = 9
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MENomPGI: TMaskEdit
            Left = 278
            Top = 121
            Width = 121
            Height = 21
            MaxLength = 35
            TabOrder = 10
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MECommentaire: TMaskEdit
            Left = 76
            Top = 169
            Width = 323
            Height = 21
            MaxLength = 35
            TabOrder = 11
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MENumero: TMaskEdit
            Left = 371
            Top = 22
            Width = 28
            Height = 21
            Color = clYellow
            MaxLength = 1
            TabOrder = 12
            Visible = False
            OnExit = OnFieldExit
          end
          object MEObligatoire: TCheckBox
            Left = 6
            Top = 99
            Width = 82
            Height = 17
            Alignment = taLeftJustify
            Caption = 'Obligatoire'
            TabOrder = 3
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object METypeAlpha: TComboBox
            Left = 278
            Top = 49
            Width = 121
            Height = 21
            Style = csDropDownList
            DropDownCount = 6
            ItemHeight = 13
            TabOrder = 7
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
            Items.Strings = (
              ''
              'Alphanum'#233'rique'
              'Num'#233'rique'
              'Date'
              'Heure')
          end
          object MEMajusMin: TComboBox
            Left = 76
            Top = 73
            Width = 95
            Height = 21
            Style = csDropDownList
            DropDownCount = 6
            ItemHeight = 13
            TabOrder = 2
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
            Items.Strings = (
              ''
              'Majuscule'
              'Minuscule')
          end
          object MEFige: TCheckBox
            Left = 119
            Top = 99
            Width = 51
            Height = 17
            Alignment = taLeftJustify
            Caption = 'Fig'#233
            TabOrder = 4
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object METableName: TMaskEdit
            Left = 76
            Top = 25
            Width = 93
            Height = 21
            MaxLength = 35
            TabOrder = 0
            OnClick = METableNameClick
            OnExit = OnFieldExit
          end
          object MENom: TMaskEdit
            Left = 278
            Top = 25
            Width = 87
            Height = 21
            MaxLength = 35
            TabOrder = 1
            OnClick = MENomClick
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MEModifiable: TCheckBox
            Left = 341
            Top = 3
            Width = 74
            Height = 17
            Alignment = taLeftJustify
            Caption = 'Modifiable'
            TabOrder = 13
            Visible = False
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MEAlignement: TComboBox
            Left = 76
            Top = 121
            Width = 95
            Height = 21
            Style = csDropDownList
            ItemHeight = 13
            TabOrder = 5
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
            Items.Strings = (
              '<<Aucun>>'
              'Gauche'
              'Droite')
          end
          object MESeparateur: TEdit
            Left = 76
            Top = 145
            Width = 42
            Height = 21
            TabOrder = 6
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MENumVersion: TMaskEdit
            Left = 383
            Top = 22
            Width = 28
            Height = 21
            Color = clYellow
            MaxLength = 1
            TabOrder = 14
            Visible = False
            OnExit = OnFieldExit
          end
          object bCalcul: THBitBtn
            Left = 373
            Top = 144
            Width = 26
            Height = 24
            TabOrder = 15
            OnClick = bCalculClick
            GlobalIndexImage = 'Z0507_S16G1'
          end
          object TEOffset: TEdit
            Left = 278
            Top = 1
            Width = 58
            Height = 21
            Enabled = False
            ReadOnly = True
            TabOrder = 16
          end
          object MELongueur: TSpinEdit
            Left = 76
            Top = 49
            Width = 73
            Height = 23
            AutoSize = False
            Enabled = False
            MaxValue = 0
            MinValue = 0
            TabOrder = 17
            Value = 1
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MENbdecimal: TSpinEdit
            Left = 278
            Top = 145
            Width = 73
            Height = 23
            AutoSize = False
            Enabled = False
            MaxValue = 0
            MinValue = 0
            TabOrder = 18
            Value = 0
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MEDomaine: THValComboBox
            Left = 75
            Top = 1
            Width = 97
            Height = 21
            Style = csDropDownList
            ItemHeight = 13
            TabOrder = 19
            OnExit = OnFieldExit
            TagDispatch = 0
            ComboWidth = 300
          end
        end
        object Page2: TTabSheet
          Caption = 'Compl'#233'ment'
          ImageIndex = 1
          object TMELibre1: TLabel
            Left = 4
            Top = 30
            Width = 64
            Height = 13
            Caption = 'Champ libre 1'
            FocusControl = MELibre1
          end
          object Label1: TLabel
            Left = 4
            Top = 54
            Width = 64
            Height = 13
            Caption = 'Champ libre 2'
            FocusControl = MELibre2
          end
          object Label2: TLabel
            Left = 4
            Top = 78
            Width = 64
            Height = 13
            Caption = 'Champ libre 3'
            FocusControl = MELibre3
          end
          object Label3: TLabel
            Left = 4
            Top = 102
            Width = 64
            Height = 13
            Caption = 'Champ libre 4'
            FocusControl = MELibre4
          end
          object Label4: TLabel
            Left = 4
            Top = 126
            Width = 64
            Height = 13
            Caption = 'Champ libre 5'
            FocusControl = MELibre5
          end
          object Label5: TLabel
            Left = 192
            Top = 30
            Width = 64
            Height = 13
            Caption = 'Champ libre 6'
            FocusControl = MELibre6
          end
          object Label6: TLabel
            Left = 192
            Top = 54
            Width = 64
            Height = 13
            Caption = 'Champ libre 7'
            FocusControl = MELibre7
          end
          object Label7: TLabel
            Left = 192
            Top = 78
            Width = 64
            Height = 13
            Caption = 'Champ libre 8'
            FocusControl = MELibre8
          end
          object Label8: TLabel
            Left = 192
            Top = 102
            Width = 64
            Height = 13
            Caption = 'Champ libre 9'
            FocusControl = MELibre9
          end
          object Label9: TLabel
            Left = 192
            Top = 126
            Width = 70
            Height = 13
            Caption = 'Champ libre 10'
            FocusControl = MELibre10
          end
          object MELibre1: TMaskEdit
            Left = 80
            Top = 26
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 0
            OnExit = OnFieldExit
          end
          object MELibre2: TMaskEdit
            Left = 80
            Top = 50
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 1
            OnExit = OnFieldExit
          end
          object MELibre3: TMaskEdit
            Left = 80
            Top = 74
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 2
            OnExit = OnFieldExit
          end
          object MELibre4: TMaskEdit
            Left = 80
            Top = 98
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 3
            OnExit = OnFieldExit
          end
          object MELibre5: TMaskEdit
            Left = 80
            Top = 122
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 4
            OnExit = OnFieldExit
          end
          object MELibre6: TMaskEdit
            Left = 268
            Top = 26
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 5
            OnExit = OnFieldExit
          end
          object MELibre7: TMaskEdit
            Left = 268
            Top = 50
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 6
            OnExit = OnFieldExit
          end
          object MELibre8: TMaskEdit
            Left = 268
            Top = 74
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 7
            OnExit = OnFieldExit
          end
          object MELibre9: TMaskEdit
            Left = 268
            Top = 98
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 8
            OnExit = OnFieldExit
          end
          object MELibre10: TMaskEdit
            Left = 268
            Top = 122
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 9
            OnExit = OnFieldExit
          end
        end
        object Page3: TTabSheet
          Caption = 'Tri'
          ImageIndex = 3
          object LTETri: TLabel
            Left = 7
            Top = 30
            Width = 52
            Height = 13
            Caption = 'Ordre de tri'
            FocusControl = MEOrdreTri
          end
          object Label13: TLabel
            Left = 7
            Top = 58
            Width = 128
            Height = 13
            Caption = 'Famille de Correspondance'
            FocusControl = MEFamCorresp
          end
          object MEFamCorresp: TMaskEdit
            Left = 140
            Top = 54
            Width = 93
            Height = 21
            MaxLength = 15
            TabOrder = 0
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
          object MEOrdreTri: TMaskEdit
            Left = 64
            Top = 26
            Width = 345
            Height = 21
            MaxLength = 255
            TabOrder = 1
            OnEnter = OnFieldEnter
            OnExit = OnFieldExit
          end
        end
      end
    end
    object SG1: THGrid
      Left = 0
      Top = 0
      Width = 438
      Height = 182
      Align = alClient
      DefaultRowHeight = 18
      FixedCols = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      TabOrder = 2
      OnClick = SG1Click
      OnEnter = SG1Enter
      OnExit = SG1Exit
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
    end
    object DBG: TDBGrid
      Left = 256
      Top = 136
      Width = 117
      Height = 53
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      Visible = False
    end
  end
  object TWCalcul: TToolWindow97
    Left = -112
    Top = 76
    ClientHeight = 213
    ClientWidth = 285
    Caption = 'Formule de calcul'
    ClientAreaHeight = 213
    ClientAreaWidth = 285
    DockableTo = []
    TabOrder = 3
    Visible = False
    OnVisibleChanged = TWCalculVisibleChanged
    object MECalcul: TMemo
      Left = 0
      Top = 0
      Width = 285
      Height = 213
      Align = alClient
      BorderStyle = bsNone
      TabOrder = 0
    end
  end
  object DS1: TDataSource
    Left = 57
    Top = 324
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
      'Z0053_S16G1'
      'Z0037_S16G1'
      'Z0241_S16G1')
    Left = 108
    Top = 375
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 108
    Top = 324
  end
  object DS3: TDataSource
    Left = 56
    Top = 376
  end
  object HMsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Champ fig'#233' : valeur obligatoire;E;O;O;O;'
      
        '1;?caption?;Longueur du texte superieure '#224' la longueur du champ;' +
        'E;O;O;O;'
      '2;?caption?;La contenu saisi n'#39'est pas num'#233'rique;E;O;O;O;'
      
        '3;?caption?;La longueur du nombre doit '#234'tre comprise entre 1 et ' +
        '7 caract'#232'res;E;O;O;O;'
      
        '4;?caption?;Voulez vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;')
    Left = 108
    Top = 276
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 56
    Top = 276
  end
end

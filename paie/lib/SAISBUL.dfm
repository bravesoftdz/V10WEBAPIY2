object FSaisBul: TFSaisBul
  Left = 250
  Top = 178
  Width = 699
  Height = 543
  HelpContext = 42211005
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie du bulletin'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  ObjectMenuItem = Cotzz
  PopupMenu = PopupMenuGlobal
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PCrit: THPanel
    Left = 0
    Top = 0
    Width = 691
    Height = 129
    Align = alTop
    Caption = ' '
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFond
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    ColorNb = 1
    TextEffect = tenone
    object Bevel1: TBevel
      Left = 1
      Top = 1
      Width = 689
      Height = 127
      Align = alClient
    end
    object Bevel4: TBevel
      Left = 3
      Top = 73
      Width = 681
      Height = 2
    end
    object Bevel3: TBevel
      Left = 440
      Top = 74
      Width = 2
      Height = 54
    end
    object LblProfilMulti: THLabel
      Left = 437
      Top = 10
      Width = 119
      Height = 13
      Caption = 'Autres Profils P'#233'riodiques'
      Visible = False
    end
    object BtnAligProfPer: TToolbarButton97
      Left = 565
      Top = 17
      Width = 28
      Height = 27
      Hint = 'Aligner le bulletin sur les profils p'#233'riodiques'
      DisplayMode = dmGlyphOnly
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Layout = blGlyphTop
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Visible = False
      OnClick = BtnAligProfPerClick
      GlobalIndexImage = 'Z0775_S16G1'
    end
    object TToolbarButton97
      Left = 655
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Acc'#232'der aux Salari'#233's'
      DisplayMode = dmGlyphOnly
      Caption = 'Salari'#233
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImageIndex = 0
      Layout = blGlyphTop
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Visible = False
      OnClick = BSalarieClick
      GlobalIndexImage = 'Z0781_S16G1'
    end
    object LblEtab: THLabel
      Left = 163
      Top = 11
      Width = 68
      Height = 13
      Caption = 'Etablissement '
      Transparent = True
    end
    object TPSA_ETABLISSEMENT: THLabel
      Left = 260
      Top = 11
      Width = 201
      Height = 13
      Hint = 'Etablissement auquel appartient le salari'#233
      AutoSize = False
      Caption = '..................................'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Transparent = True
    end
    object TSalarie: THLabel
      Left = 163
      Top = 27
      Width = 32
      Height = 13
      Caption = 'Salari'#233
      Transparent = True
    end
    object FSalarie: THLabel
      Left = 260
      Top = 27
      Width = 74
      Height = 13
      Hint = 'Matricule du salari'#233
      AutoSize = False
      Caption = '......'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Transparent = True
    end
    object TPSA_LIBELLE: THLabel
      Left = 336
      Top = 27
      Width = 328
      Height = 13
      Hint = 'Nom du salari'#233
      AutoSize = False
      Caption = '..................................'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Transparent = True
    end
    object TDateEntree: THLabel
      Left = 163
      Top = 43
      Width = 64
      Height = 13
      Caption = 'Date d'#39'entr'#233'e'
      Transparent = True
    end
    object LblDateSortie: THLabel
      Left = 328
      Top = 43
      Width = 66
      Height = 13
      Caption = 'Date de sortie'
      Transparent = True
    end
    object TPSA_DATEENTREE: THLabel
      Left = 260
      Top = 43
      Width = 60
      Height = 13
      AutoSize = False
      Caption = '....'
      Transparent = True
    end
    object LDateNaissance: THLabel
      Left = 163
      Top = 57
      Width = 89
      Height = 13
      Caption = 'Date de naissance'
      Transparent = True
    end
    object TPSA_DATENAISS: THLabel
      Left = 260
      Top = 57
      Width = 60
      Height = 13
      AutoSize = False
      Caption = '....'
      Transparent = True
    end
    object TPSA_DATESORTIE: THLabel
      Left = 404
      Top = 43
      Width = 60
      Height = 13
      AutoSize = False
      Caption = '....'
      Transparent = True
    end
    object LTitre: TLabel
      Left = 4
      Top = 13
      Width = 153
      Height = 23
      Alignment = taCenter
      AutoSize = False
      BiDiMode = bdLeftToRight
      Caption = 'BULLETIN'
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'Arial Black'
      Font.Style = [fsBold]
      ParentBiDiMode = False
      ParentColor = False
      ParentFont = False
      ParentShowHint = False
      ShowAccelChar = False
      ShowHint = False
      Transparent = True
    end
    object FTypeActionBul: TLabel
      Left = 4
      Top = 42
      Width = 153
      Height = 20
      Alignment = taCenter
      AutoSize = False
      Caption = '..............'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      Transparent = True
    end
    object LblPeriodePaie: THLabel
      Left = 5
      Top = 81
      Width = 14
      Height = 13
      Caption = 'Du'
      Transparent = True
    end
    object HLabel3: THLabel
      Left = 135
      Top = 81
      Width = 12
      Height = 13
      Caption = 'au'
      Transparent = True
    end
    object LblSlash: TLabel
      Left = 382
      Top = 80
      Width = 5
      Height = 13
      Caption = '/'
      Transparent = True
    end
    object HLabel1: THLabel
      Left = 545
      Top = 81
      Width = 32
      Height = 13
      Caption = 'Acquis'
      Transparent = True
    end
    object TCPSUPP: THLabel
      Left = 594
      Top = 81
      Width = 30
      Height = 13
      Caption = 'Suppl.'
      Transparent = True
    end
    object TCPANC: THLabel
      Left = 646
      Top = 81
      Width = 22
      Height = 13
      Caption = 'Anc.'
      Transparent = True
    end
    object FrancImage: TImage
      Left = 662
      Top = 10
      Width = 20
      Height = 17
      Picture.Data = {
        07544269746D617036010000424D360100000000000076000000280000001400
        0000100000000100040000000000C0000000C40E0000C40E0000100000000000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00777777777777777777770000777777777777777777770000747777747774
        7747777700007477777477747747777700007477777477747747777700007477
        7774774477477777000074777774774777477777000074777774444777477777
        0000747777744444774777770000744447747774774444770000747777747774
        7747777700007477777477747747777700007477777477747747777700007477
        7774777477477777000074444474444477444447000077777777777777777777
        0000}
      Visible = False
    end
    object EuroImage: TImage
      Left = 660
      Top = 10
      Width = 16
      Height = 17
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        000010000000010004000000000080000000C40E0000C40E0000100000000000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00777777777777777777777774444444477777744444444444777744447777
        7774777744477777777777774477777777777444444444447777774444444444
        4777777447777777777774444444444447777744444444444477777744777777
        7774777744477777774477777444777774447777774444444444777777744444
        4774}
      Visible = False
    end
    object LBLEDTDU: THLabel
      Left = 5
      Top = 104
      Width = 39
      Height = 13
      Caption = 'Edit'#233' du'
      Transparent = True
    end
    object LBLEDTAU: THLabel
      Left = 135
      Top = 104
      Width = 12
      Height = 13
      Caption = 'au'
      Transparent = True
    end
    object Image1: TImage
      Left = 660
      Top = 10
      Width = 16
      Height = 17
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        000010000000010004000000000080000000C40E0000C40E0000100000000000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00777777777777777777777774444444477777744444444444777744447777
        7774777744477777777777774477777777777444444444447777774444444444
        4777777447777777777774444444444447777744444444444477777744777777
        7774777744477777774477777444777774447777774444444444777777744444
        4774}
      Visible = False
    end
    object ImageBulCompl: TImage
      Left = 626
      Top = 30
      Width = 16
      Height = 17
      Picture.Data = {
        07544269746D617036010000424D360100000000000076000000280000001200
        0000100000000100040000000000C00000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFF000000FFFFF00000000FFFFF000000FFFF87B77088
        800FFF000000FFFF8FFFBF0B7880FF000000FFFFF8BFFFB0FB780F000000FFFF
        F8FFBFF0F0080F000000FFFFFF8FFFB80FB70F000000FFFFFF00BFFFBFF70F00
        0000FF00F0660FBFFFB70F000000F06600060FFFBFF70F000000F0E60E06000F
        FFB70F000000F0EE00EE0FFFBFF70F000000FF00FF00FFBFFFB70F000000FFFF
        FF8FBFFFBFF70F000000FFFFFF88888888880F000000FFFFFFFFFFFFFFFFFF00
        0000}
      Transparent = True
    end
    object ImageCalcul: TImage
      Left = 657
      Top = 38
      Width = 20
      Height = 17
      Picture.Data = {
        07544269746D617076050000424D760500000000000036040000280000001400
        0000100000000100080000000000400100000000000000000000000100000001
        0000000000007B0000007B7B0000FFFF0000BDBDBD00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00000000000000BF0000BF000000BF
        BF00BF000000BF00BF00BFBF0000C0C0C000808080000000FF0000FF000000FF
        FF00FF000000FF00FF00FFFF0000FFFFFF000000000000000000000000000000
        0000F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9
        F9F9F9F9F9F9F9F9F9F9F9F9F9FFFFFFFFFFFFFFFFFFFFFFFFFFFFF9F9F9F9F9
        020201010101010101010101010101FFF9F9F9F902FB02020202020202020202
        020201FFF9F9F9F902FA02FFFBFFFBFFFBFFFBFBF3FF01FFF9F9F9F902FB0202
        0202020202020202020201FFF9F9F9F902FA02FFFBFFFBFFFBFFFBFFFBFF01FF
        F9F9F9F902FB02020202020202020202020201FFF9F9F9F902FAFFF3F3F3F3F3
        F3F30202020201FFF9F9F9F902FBFFFBFBFBFBFBFBF30202020201FFF9F9F9F9
        02FAFFFFFFFFFFFFFFFF0202020201FFF9F9F9F902FBFAFBFAFBFAFBFAFBFAFB
        FAFB02FFF9F9F9F9F90202020202020202020202020202F9F9F9F9F9F9F9F9F9
        F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9
        F9F9}
      Transparent = True
    end
    object LbllTrentieme: THLabel
      Left = 258
      Top = 80
      Width = 74
      Height = 13
      Caption = 'Trenti'#232'me for'#231#233
      Transparent = True
    end
    object HFSortieDef: THLabel
      Left = 511
      Top = 43
      Width = 72
      Height = 13
      Caption = 'Sortie d'#233'finitive'
      Transparent = True
    end
    object HCBModifAcquis: THLabel
      Left = 463
      Top = 104
      Width = 73
      Height = 13
      Caption = 'Acquis modifi'#233's'
      Transparent = True
    end
    object Bevel5: TBevel
      Left = 235
      Top = 74
      Width = 2
      Height = 54
    end
    object Bevel6: TBevel
      Left = 156
      Top = 1
      Width = 2
      Height = 73
    end
    object FSortieDef: TCheckBox
      Left = 491
      Top = 41
      Width = 13
      Height = 17
      Color = clBtnFace
      Enabled = False
      ParentColor = False
      TabOrder = 9
    end
    object MulProfilTempo: THMultiValComboBox
      Left = 499
      Top = 8
      Width = 71
      Height = 21
      Hint = 'S'#233'lectionner les profils '#224' appliquer sur CE bulletin'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 14
      Visible = False
      Abrege = False
      DataType = 'PGPROFILDIV'
      Complete = False
      OuInclusif = False
    end
    object FDate1: THCritMaskEdit
      Left = 49
      Top = 77
      Width = 81
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 0
      Text = '  /  /    '
      OnEnter = FDatePaieEnter
      OnExit = FDate1Exit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      DefaultDate = odDebMois
      ElipsisButton = True
      ControlerDate = True
    end
    object FDate2: THCritMaskEdit
      Left = 152
      Top = 77
      Width = 81
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '  /  /    '
      OnEnter = FDatePaieEnter
      OnExit = FDate2Exit
      TagDispatch = 0
      Operateur = Inferieur
      OpeType = otDate
      DefaultDate = odFinMois
      ElipsisButton = True
      ControlerDate = True
    end
    object FTrentieme: THCritMaskEdit
      Left = 351
      Top = 76
      Width = 28
      Height = 21
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnEnter = FTrentiemeEnter
      OnExit = FTrentiemeExit
      TagDispatch = 0
      Operateur = Egal
    end
    object FTrentDenominateur: THCritMaskEdit
      Left = 389
      Top = 76
      Width = 28
      Height = 21
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Text = '30'
      OnEnter = FTrentDenominateurEnter
      OnExit = FTrentDenominateurExit
      TagDispatch = 0
      Operateur = Egal
    end
    object ChbxTrentiem: TCheckBox
      Left = 240
      Top = 78
      Width = 13
      Height = 17
      TabOrder = 4
      OnClick = ChbxTrentiemeClick
    end
    object CPMOIS: THCritMaskEdit
      Left = 538
      Top = 100
      Width = 45
      Height = 21
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 11
      OnExit = CPAcquisExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otReel
      DisplayFormat = '0.00'
    end
    object CPSUPP: THCritMaskEdit
      Left = 586
      Top = 100
      Width = 45
      Height = 21
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 12
      OnExit = CPAcquisExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otReel
      DisplayFormat = '0.00'
    end
    object CPANC: THCritMaskEdit
      Left = 635
      Top = 100
      Width = 45
      Height = 21
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 13
      OnExit = CPAcquisExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otReel
      ControlerDate = True
      DisplayFormat = '0.00'
    end
    object FEDTDU: THCritMaskEdit
      Left = 49
      Top = 100
      Width = 81
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 2
      Text = '  /  /    '
      OnExit = FDateDuExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = True
    end
    object FEDTAU: THCritMaskEdit
      Left = 152
      Top = 100
      Width = 81
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 3
      Text = '  /  /    '
      OnExit = FDateAuExit
      TagDispatch = 0
      Operateur = Inferieur
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = True
    end
    object CBModifAcquis: TCheckBox
      Left = 447
      Top = 102
      Width = 13
      Height = 17
      Color = clBtnFace
      ParentColor = False
      TabOrder = 10
      OnClick = CBModifAcquisClick
    end
    object ChbxBaseForcee: TCheckBox
      Left = 240
      Top = 102
      Width = 93
      Height = 17
      Caption = 'Bases forc'#233'es'
      TabOrder = 7
      OnClick = ZChbxBaseForceeClick
    end
    object ChbxTranchesForcees: TCheckBox
      Left = 332
      Top = 102
      Width = 107
      Height = 17
      Caption = 'Tranches forc'#233'es'
      TabOrder = 8
      OnClick = ZChbxTranchesForceesClick
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 460
    Width = 691
    Height = 49
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 45
      ClientWidth = 691
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 45
      ClientAreaWidth = 691
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object Binsert: TToolbarButton97
        Left = 144
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Caption = 'Nouveau'
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
        OnClick = BinsertClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 570
        Top = 4
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
        Left = 600
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Valider'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
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
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object bDefaire: TToolbarButton97
        Left = 116
        Top = 4
        Width = 28
        Height = 27
        Hint = 
          ' R'#233'alignement des rubriques du bulletin par rapport aux profils ' +
          'du salari'#233
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 4
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bDefaireClick
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object HelpBtn: TToolbarButton97
        Left = 661
        Top = 4
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
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
          A600000000000000000000000000000000000000000000000000000000000000
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
        GlyphMask.Data = {00000000}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = HelpBtnClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 631
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
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
      object BDelete: TToolbarButton97
        Left = 172
        Top = 4
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
      object bNewligne: TToolbarButton97
        Left = 201
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une ligne'
        DisplayMode = dmGlyphOnly
        Caption = 'Ins'#233'rer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bNewligneClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object bDelLigne: TToolbarButton97
        Left = 230
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Supprimer la ligne en cours'
        DisplayMode = dmGlyphOnly
        Caption = 'Effacer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bDelLigneClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 492
        Top = 28
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la pi'#232'ce'
        DisplayMode = dmGlyphOnly
        Caption = 'Rechercher'
        Enabled = False
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BCalculBull: TToolbarButton97
        Left = 537
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Calcul du bulletin'
        DisplayMode = dmGlyphOnly
        Caption = 'Calcul Bulletin'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BCalculBullClick
        GlobalIndexImage = 'Z0209_S16G1'
      end
      object BtnDetail: TToolbarButton97
        Left = 416
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Pr'#233'sentation d'#233'taill'#233'e/simplifi'#233'e du bulletin'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BtnDetailClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BtnSimple: TToolbarButton97
        Left = 621
        Top = 28
        Width = 28
        Height = 27
        Hint = 'Pr'#233'sentation simplifi'#233'e du bulletin'
        DisplayMode = dmGlyphOnly
        Enabled = False
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BtnSimpleClick
        GlobalIndexImage = 'Z0210_S16G1'
      end
      object BFirst: TToolbarButton97
        Left = 2
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Premier'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BFirstClick
        GlobalIndexImage = 'Z0301_S16G1'
      end
      object BPrev: TToolbarButton97
        Left = 30
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Pr'#233'c'#233'dent'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BPrevClick
        GlobalIndexImage = 'Z0017_S16G1'
      end
      object BNext: TToolbarButton97
        Left = 58
        Top = 4
        Width = 31
        Height = 27
        Hint = 'Suivant'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BNextClick
        GlobalIndexImage = 'Z0031_S16G1'
      end
      object BLast: TToolbarButton97
        Left = 89
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Dernier'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BLastClick
        GlobalIndexImage = 'Z0264_S16G1'
      end
      object BCommentaire: TToolbarButton97
        Left = 293
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une ligne de commentaire'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = LigneCommentaire
        GlobalIndexImage = 'Z1827_S16G0'
      end
      object BZoom: TToolbarButton97
        Left = 259
        Top = 4
        Width = 34
        Height = 27
        Hint = 'Zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopupMenuZoom
        Caption = 'Zoom'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BZoomClick
        GlobalIndexImage = 'Z0061_S16G1'
      end
      object BRechargContext: TToolbarButton97
        Left = 472
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Recharger Contexte de la Paie'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechargContextClick
        GlobalIndexImage = 'M0024_S16G1'
      end
      object BtnOrigine: TToolbarButton97
        Left = 352
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Param'#233'trage de l'#39'origine de la rubrique'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BtnOrigineClick
        GlobalIndexImage = 'Z0008_S16G1'
      end
      object BVentil: TToolbarButton97
        Left = 382
        Top = 4
        Width = 34
        Height = 27
        Hint = 'R'#233'partition analytique'
        DropdownArrow = True
        DropdownMenu = PopupMenuAnal
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0133_S16G1'
      end
      object BRegul: TToolbarButton97
        Left = 323
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une ligne de r'#233'gularisation'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = LigneRegul
        GlobalIndexImage = 'Z0775_S16G1'
      end
      object BTNDIAGNOSTIC: TToolbarButton97
        Left = 502
        Top = 4
        Width = 34
        Height = 27
        Hint = 'Outil de diagnostic'
        DropdownArrow = True
        DropdownMenu = PopupMenuDiag
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0298_S16G1'
      end
      object BtnSaisie: TToolbarButton97
        Left = 444
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Voir uniquement les r'#233'mun'#233'rations saisissables'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BtnDetailClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
    end
  end
  object PTotaux: THPanel
    Left = 0
    Top = 408
    Width = 691
    Height = 52
    Align = alBottom
    Caption = ' '
    Ctl3D = True
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFond
    ColorShadow = clBtnText
    ColorStart = clBtnFace
    ColorNb = 1
    TextEffect = tenone
    object BlvlCumul: TBevel
      Left = 1
      Top = 4
      Width = 459
      Height = 45
    end
    object LblCumBrut: THLabel
      Left = 10
      Top = 10
      Width = 19
      Height = 13
      Caption = 'Brut'
      Transparent = True
    end
    object LblCumBrutFiscal: THLabel
      Left = 10
      Top = 28
      Width = 46
      Height = 13
      Caption = 'Brut fiscal'
      Transparent = True
    end
    object LblCumNetApayer: THLabel
      Left = 170
      Top = 28
      Width = 55
      Height = 13
      Caption = 'Net '#224' payer'
      Transparent = True
    end
    object LblCumNetImp: THLabel
      Left = 170
      Top = 10
      Width = 67
      Height = 13
      Caption = 'Net imposable'
      Transparent = True
    end
    object LblCumHeures: THLabel
      Left = 338
      Top = 10
      Width = 59
      Height = 13
      Caption = 'Total heures'
      Transparent = True
    end
    object LblBrut: THLabel
      Left = 60
      Top = 10
      Width = 101
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '0.00'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object LblBrutFiscal: THLabel
      Left = 60
      Top = 28
      Width = 101
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '0.00'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object LblNetImp: THLabel
      Left = 240
      Top = 10
      Width = 91
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '0.00'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object LblNetAPayer: THLabel
      Left = 240
      Top = 28
      Width = 91
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '0.00'
      Color = clAqua
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      Transparent = True
    end
    object LblHeures: THLabel
      Left = 416
      Top = 10
      Width = 41
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '0.00'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object Bevel2: TBevel
      Left = 459
      Top = 4
      Width = 222
      Height = 45
    end
    object LblPaye: THLabel
      Left = 462
      Top = 28
      Width = 35
      Height = 13
      Caption = 'Pay'#233' le'
      Transparent = True
    end
    object LblPar: THLabel
      Left = 588
      Top = 28
      Width = 15
      Height = 13
      Caption = 'par'
      Transparent = True
    end
    object HLabel6: THLabel
      Left = 499
      Top = 8
      Width = 87
      Height = 13
      Caption = 'R'#232'glement modifi'#233
      Transparent = True
    end
    object ModeRegle: THValComboBox
      Left = 606
      Top = 25
      Width = 73
      Height = 21
      Enabled = False
      ItemHeight = 13
      TabOrder = 1
      OnExit = ModeRegleOnExit
      TagDispatch = 0
      DataType = 'PGMODEREGLE'
    end
    object FDatePaie: THCritMaskEdit
      Left = 499
      Top = 24
      Width = 86
      Height = 21
      Enabled = False
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 0
      Text = '  /  /    '
      OnExit = FDatePaieOnExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = True
    end
    object ChbxRegltMod: TCheckBox
      Left = 463
      Top = 8
      Width = 14
      Height = 17
      TabOrder = 2
      OnClick = ChbxRegltModClick
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 129
    Width = 691
    Height = 279
    ActivePage = PRemNonImp
    Align = alClient
    TabOrder = 3
    OnChange = PagesChange
    object PRemSais: TTabSheet
      Caption = 'El'#233'ments '#224' saisir'
      ImageIndex = 1
      TabVisible = False
      object GRemSais: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          48
          209
          87
          70
          74
          98
          68)
      end
    end
    object PRemSal: TTabSheet
      Caption = 'Salaires'
      object GRemSal: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          48
          209
          87
          70
          74
          98
          68)
      end
    end
    object PRemHeures: TTabSheet
      Caption = 'Heures'
      ImageIndex = 3
      object GRemHeures: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          98
          70
          74
          90
          51)
      end
    end
    object PRemPrimes: TTabSheet
      Caption = 'Primes'
      ImageIndex = 4
      object GRemPrimes: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 255
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C'
          ' ')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          98
          70
          74
          92
          47)
      end
    end
    object PRemAbs: TTabSheet
      Caption = 'Absences'
      ImageIndex = 5
      object LblAbsHeures: TLabel
        Left = 153
        Top = 8
        Width = 57
        Height = 13
        Caption = 'Horaire r'#233'el '
      end
      object LblOuvres: TLabel
        Left = 273
        Top = 8
        Width = 97
        Height = 13
        Caption = 'Nombre jours ouvr'#233's'
      end
      object LblOuvrable: TLabel
        Left = 457
        Top = 8
        Width = 111
        Height = 13
        Caption = 'Nombre jours ouvrables'
      end
      object GRemAbs: THGrid
        Left = 12
        Top = 32
        Width = 629
        Height = 169
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          92
          63
          67
          88
          48)
      end
      object ChbxHoraireSalarie: TCheckBox
        Left = 10
        Top = 6
        Width = 129
        Height = 17
        Caption = 'Horaire personnalis'#233
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = ChbxHoraireSalarieClick
      end
      object EdtHorReel: TEdit
        Left = 224
        Top = 4
        Width = 41
        Height = 21
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnExit = EdtHorReelExit
      end
      object EdtJoursOuvres: TEdit
        Left = 390
        Top = 3
        Width = 48
        Height = 21
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnExit = EdtJoursOuvresExit
      end
      object EdtJoursOuvrables: TEdit
        Left = 594
        Top = 3
        Width = 47
        Height = 21
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnExit = EdtJoursOuvrablesExit
      end
    end
    object PRemCplt: TTabSheet
      Caption = 'Compl'#233'ments'
      ImageIndex = 6
      object GRemCplt: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 255
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          98
          70
          74
          90
          50)
      end
    end
    object PRemAvt: TTabSheet
      Caption = 'Avantages'
      ImageIndex = 7
      object GRemAvt: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 259
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          98
          70
          74
          90
          50)
      end
    end
    object PRemAbt: TTabSheet
      Caption = 'Abattements'
      ImageIndex = 8
      object GRemAbt: THGrid
        Left = 8
        Top = 32
        Width = 633
        Height = 161
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          91
          70
          70
          87
          46)
      end
      object ChbxControlSmic: TCheckBox
        Left = 16
        Top = 4
        Width = 117
        Height = 17
        Caption = 'Contr'#244'le limite SMIC'
        TabOrder = 1
      end
    end
    object PRemPrimeNonImpSoumis: TTabSheet
      Caption = 'Primes NI soumises Cotisat'
      ImageIndex = 11
      object GRemNonImpSoumis: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 255
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          98
          70
          74
          90
          50)
      end
    end
    object PBas: TTabSheet
      Caption = 'Bases de cotisation'
      ImageIndex = 1
      object GBas: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 255
        Align = alClient
        ColCount = 11
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 9
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GBaseDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Plafond;D'
          'Tranche 1;D'
          'Tranche 2;D'
          'Tranche 3;D'
          'Plafond Tr1;D'
          'Plafond Tr2;D'
          'Plafond Tr3;D'
          'Origine;G')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GBasRowEnter
        OnRowExit = GBasRowExit
        OnCellEnter = GBaseSalCellEnter
        OnCellExit = GBaseSalCellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GCotElipsisClick
        ColWidths = (
          43
          206
          78
          73
          81
          69
          72
          63
          64
          58
          64)
      end
      object ZChbxBaseForcee: TCheckBox
        Left = 12
        Top = 1
        Width = 97
        Height = 17
        Caption = 'Bases Forc'#233'es'
        TabOrder = 1
        Visible = False
        OnClick = ZChbxBaseForceeClick
      end
      object ZChbxTranchesForcees: TCheckBox
        Left = 124
        Top = 1
        Width = 109
        Height = 17
        Caption = 'Tranches Forc'#233'es'
        TabOrder = 2
        Visible = False
        OnClick = ZChbxTranchesForceesClick
      end
    end
    object PCot: TTabSheet
      Caption = 'Cotisations'
      ImageIndex = 2
      object GCot: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        ColCount = 8
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 15
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GCotDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Tx. Sal.;D'
          'Montant Sal;D'
          'Tx. Pat.;D'
          'Montant Pat.;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GCotRowEnter
        OnRowExit = GCotRowExit
        OnCellEnter = GCotSalCellEnter
        OnCellExit = GCotSalCellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GCotElipsisClick
        ColWidths = (
          44
          188
          73
          52
          79
          50
          79
          58)
      end
    end
    object PRemRet: TTabSheet
      Caption = 'Retenues'
      ImageIndex = 9
      object GRemRet: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          213
          98
          70
          74
          90
          49)
      end
    end
    object PRemNonImp: TTabSheet
      Caption = 'Primes Non Imposables'
      ImageIndex = 10
      object GRemNonImp: THGrid
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        ColCount = 7
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnDblClick = GRemSalDblClick
        OnEnter = OnGrilleEnter
        OnExit = OnGrilleExit
        OnKeyDown = GRemKeyDown
        OnSetEditText = GRemSalSetEditText
        SortedCol = -1
        Titres.Strings = (
          'Code;G'
          'Libell'#233';G'
          'Base;D'
          'Taux;D'
          'Coefficient;D'
          'Montant;D'
          'Origine;C')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GRemSalRowEnter
        OnRowExit = GRemSalRowExit
        OnCellEnter = GRemSalCellEnter
        OnCellExit = GRemSalCellExit
        ColCombo = 6
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        OnElipsisClick = GRemSalElipsisClick
        ColWidths = (
          45
          216
          98
          70
          74
          91
          46)
      end
    end
    object TBSHTDIAG: TTabSheet
      Caption = 'Diagnostic'
      ImageIndex = 12
      object Diag: TTreeView
        Left = 0
        Top = 0
        Width = 683
        Height = 251
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Indent = 19
        MultiSelect = True
        ParentFont = False
        TabOrder = 0
      end
    end
  end
  object HMTrad: THSystemMenu
    MaxInsideSize = False
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 144
    Top = 288
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      
        '1;?caption?;Confirmez-vous la suppression du bulletin ?;Q;YNC;N;' +
        'C;'
      
        '2;?caption?;Confirmez-vous le retour '#224' l'#39#233'tat initial du bulleti' +
        'n?;Q;YNC;N;C;'
      
        '3;?caption?;Retour Etat initial et Alignement du bulletin avec l' +
        #39'ensemble des profils?;Q;YN;Y;N;')
    Left = 60
    Top = 287
  end
  object ImageList1: THImageList
    GlobalIndexImages.Strings = (
      'Z0223_S16G1'
      'Z1831_S16G1'
      'Z1374_S16G1'
      'Z1381_S16G1'
      'Z1862_S16G1'
      'Z1952_S16G1'
      'Z1995_S16G1'
      'Z2124_S16G1')
    Height = 18
    Width = 18
    Left = 376
    Top = 287
  end
  object PopupMenuGlobal: TPopupMenu
    Images = ImageList1
    Left = 232
    Top = 292
    object PaieEnvers: TMenuItem
      Caption = 'Calcul de la &paie '#224' l'#39'envers'
      OnClick = PaieEnversClick
    end
    object PaieNetRub: TMenuItem
      Caption = 'Net '#224' payer '#224' partir d'#39'une rubrique'
      OnClick = PaieEnversRubClick
    end
    object Zoom: TMenuItem
      Caption = 'Zoom (Info concernant le salari'#233')'
      OnClick = BZoomClick
    end
    object EtablissPaie: TMenuItem
      Caption = '&Etablissement'
      Hint = 'Acc'#232'der '#224' l'#39'Etablissement'
      RadioItem = True
      OnClick = EtablissPaieClick
    end
    object Sal: TMenuItem
      Caption = '&Salari'#233's '
      Hint = 'Acc'#232'der aux Salari'#233's'
      OnClick = SalClick
    end
    object Rem: TMenuItem
      Caption = '&R'#233'mun'#233'rations'
      Hint = 'Acc'#232'der aux R'#233'mun'#233'rations'
      OnClick = RemClick
    end
    object CumSess: TMenuItem
      Caption = 'C&umuls '
      Hint = 'Acc'#232'der aux Cumuls En Cours'
      OnClick = CumSessClick
    end
    object Cotzz: TMenuItem
      Caption = '&Cotisations '
      Enabled = False
      Hint = 'Acc'#232'der aux Cotisations et aux Bases de cotisation'
      Visible = False
      OnClick = CotzzClick
    end
    object Cotisations1: TMenuItem
      Caption = '&Cotisations '
      Hint = 'Acc'#232'der aux Cotisations et aux Bases de cotisation'
      OnClick = CotzzClick
    end
    object Variab: TMenuItem
      Caption = '&Variables'
      Hint = 'Acc'#232'der aux variables'
      OnClick = VariabClick
    end
    object ProfilPgi: TMenuItem
      Caption = '&Profils'
      Hint = 'Acc'#232'der aux profils'
      OnClick = ProfilPgiClick
    end
    object ProfilRub: TMenuItem
      Caption = '&Liste des profils rubriques'
      Hint = 'Liste des profils contenant une rubrique'
      OnClick = ProfilRubClick
    end
    object Tablesdossier: TMenuItem
      Caption = '&Tables dossier'
      Hint = 'Liste des tables du dossier'
      OnClick = TablesdossierClick
    end
    object Tablesdynamiques1: TMenuItem
      Caption = 'Tables dynamiques'
      Hint = 'Liste des tables dynamiques'
      OnClick = TablesdynamiquesClick
    end
    object EltNationaux: TMenuItem
      Caption = 'El'#233'ments nationaux'
      OnClick = EltNationauxClick
    end
    object VisualiserBulletins: TMenuItem
      Caption = 'Visualiser les bulletins du salari'#233
      OnClick = VisualiserBulletinsClick
    end
    object CalculdesIJSS1: TMenuItem
      Caption = 'Calcul des IJSS'
      OnClick = CalculIJSSClick
    end
    object Editiondudiagnostic1: TMenuItem
      Caption = 'Edition du diagnostic'
      OnClick = Editiondudiagnostic1Click
    end
  end
  object PopupMenuZoom: TPopupMenu
    Images = ImageList1
    Left = 484
    Top = 293
    object Salarie1: TMenuItem
      Caption = '&Salari'#233
      ShortCut = 16467
      OnClick = SalarieClick
    end
    object Etablissement1: TMenuItem
      Caption = '&Etablissement'
      ShortCut = 16453
      OnClick = EtablissementClick
    end
    object CumSessSal1: TMenuItem
      Caption = 'C&umul'
      ShortCut = 16469
      OnClick = CumSessSalClick
    end
    object FicheInd1: TMenuItem
      Caption = '&Fiche individuelle'
      ShortCut = 16454
      OnClick = FicheIndClick
    end
    object VoirhistoriqueRemunrations1: TMenuItem
      Caption = 'Historique &r'#233'mun'#233'rations'
      ShortCut = 16466
      OnClick = VoirhistoriqueRemunrations1Click
    end
    object BasesCotis1: TMenuItem
      Caption = 'Historique &bases de cotisations'
      ShortCut = 16450
      OnClick = BasesCotis1Click
    end
    object VoirhistoriquecoTisations1: TMenuItem
      Caption = 'Historique co&tisations'
      ShortCut = 16468
      OnClick = VoirhistoriquecoTisations1Click
    end
    object Voirhistoriquecumuls1: TMenuItem
      Caption = 'Historique cumuls'
      OnClick = Voirhistoriquecumuls1Click
    end
    object Calendrier: TMenuItem
      Caption = 'Ca&lendrier'
      ShortCut = 16460
      OnClick = CalendrierClick
    end
    object VoirCP: TMenuItem
      Caption = '&Cong'#233's pay'#233's'
      object CongesPayes1: TMenuItem
        Caption = '&Gestion des cong'#233's pay'#233's'
        ShortCut = 16455
        OnClick = CongesPayesClick
      end
      object VoirvalorisationAbsence1: TMenuItem
        Caption = 'Valorisation &absence cong'#233's pay'#233's'
        ShortCut = 16449
        OnClick = ValoAbsenceClick
      end
      object Voir1: TMenuItem
        Caption = '&Indemnit'#233' cong'#233's pay'#233's'
        ShortCut = 16457
        OnClick = IndemcpClick
      end
      object Voirindemnitcompensatricedecongs1: TMenuItem
        Caption = 'Indemnit'#233' com&pens. cong'#233's pay'#233's'
        ShortCut = 16464
        OnClick = IndemSoldeClick
      end
      object CalculCpPaie: TMenuItem
        Caption = 'Recalcul des cong'#233's pay'#233's'
        OnClick = CalculCpPaieClick
      end
    end
    object Absence: TMenuItem
      Caption = 'Abse&nce'
      ShortCut = 16462
      OnClick = AbsenceClick
    end
    object VSaisieArret: TMenuItem
      Caption = 'Simulation saisie arr'#234't'
      ShortCut = 16471
      OnClick = VSaisieArretClick
    end
  end
  object TimerBul: TTimer
    OnTimer = TimerBulTimer
    Left = 316
    Top = 285
  end
  object PopupMenuDiag: TPopupMenu
    Images = ImageList1
    Left = 460
    Top = 253
    object DiagRub: TMenuItem
      Caption = 'Diagnostic rubrique'
      OnClick = DiagRubClick
    end
    object DiagBull: TMenuItem
      Caption = 'Diagnostic bulletin'
      OnClick = DiagBullClick
    end
    object SauvDiag: TMenuItem
      Caption = 'Sauvegarder le diagnostic'
      OnClick = SauvDiagClick
    end
    object EnvMailDiag: TMenuItem
      Caption = 'Sauvegarder et envoyer le diagnostic par mail'
      OnClick = EnvMailDiagClick
    end
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = 'log'
    FileName = 'diagnostic'
    Filter = '*.log|*.log|*.txt|*.txt|*.*|*.*'
    Left = 588
    Top = 345
  end
  object PopupMenuAnal: TPopupMenu
    Left = 436
    Top = 337
    object AnalSal: TMenuItem
      Caption = 'Analytique salari'#233
      OnClick = AnalSalClick
    end
    object Analrubbul: TMenuItem
      Caption = 'Analytique rubrique/bulletin'
      OnClick = AnalrubbulClick
    end
  end
end

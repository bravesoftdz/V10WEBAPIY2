object FDocRegl: TFDocRegl
  Left = 267
  Top = 191
  HelpContext = 7505600
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Impression des lettres-ch'#232'que'
  ClientHeight = 383
  ClientWidth = 644
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pBoutons: TPanel
    Left = 0
    Top = 348
    Width = 644
    Height = 35
    Align = alBottom
    BevelOuter = bvLowered
    TabOrder = 0
    object Action: TLabel
      Left = 8
      Top = 10
      Width = 37
      Height = 13
      Caption = 'Action'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object bAnnuler: THBitBtn
      Left = 556
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = bAnnulerClick
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
      NumGlyphs = 2
    end
    object BAide: THBitBtn
      Left = 588
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Aide'
      HelpContext = 1395200
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BImprimer: THBitBtn
      Left = 301
      Top = 3
      Width = 153
      Height = 27
      Hint = 'Imprimer les lettres ch'#232'ques s'#233'lectionn'#233'es'
      Caption = 'Imprimer les ch'#232'ques'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BImprimerClick
      Margin = 3
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BValider: THBitBtn
      Left = 524
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Valider'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BValiderClick
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
      Margin = 3
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object bApercu: THBitBtn
      Left = 460
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Aper'#231'u avant impression'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = bApercuClick
      Margin = 4
      GlobalIndexImage = 'Z0056_S16G1'
    end
    object bTest: THBitBtn
      Left = 492
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Test de cadrage'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = bTestClick
      Margin = 4
      GlobalIndexImage = 'Z1360_S16G1'
    end
  end
  object P: TPanel
    Left = 0
    Top = 79
    Width = 294
    Height = 269
    Align = alClient
    BevelOuter = bvNone
    Caption = ' '
    TabOrder = 1
    object GE: THGrid
      Tag = 1
      Left = 0
      Top = 0
      Width = 294
      Height = 269
      Align = alClient
      ColCount = 6
      Ctl3D = True
      DefaultColWidth = 50
      DefaultRowHeight = 18
      FixedCols = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      ParentCtl3D = False
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Tiers;G;S'
        'Pi'#232'ce;D;N'
        'Ech'#233'ance;C;D'
        'Montant;D;R'
        'Devise;D;R'
        'Collectif;G;S;')
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
        54
        43
        64
        78
        50
        64)
    end
  end
  object P2: TPanel
    Left = 328
    Top = 79
    Width = 316
    Height = 269
    Align = alRight
    BevelOuter = bvLowered
    Caption = ' '
    TabOrder = 2
    object G: THGrid
      Left = 1
      Top = 1
      Width = 314
      Height = 267
      Align = alClient
      ColCount = 9
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Tiers;G;S'
        'Libell'#233';G;S'
        'Montant;D;R'
        'Ch'#232'que;C;S'
        'Origines;G;S'
        'BanquePrevi;G;S;'
        'Devise;D;R'
        'Ech'#233'ances;D;D;'
        'Collectif;G;S')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      OnSorted = GSorted
      ColWidths = (
        58
        98
        75
        52
        64
        64
        64
        64
        64)
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 644
    Height = 56
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 3
    object HLabel4: THLabel
      Left = 11
      Top = 8
      Width = 90
      Height = 13
      Caption = 'Compte de banque'
      FocusControl = CompteBanque
    end
    object HLabel5: THLabel
      Left = 327
      Top = 6
      Width = 37
      Height = 13
      Caption = 'Banque'
    end
    object HLabel6: THLabel
      Left = 328
      Top = 34
      Width = 35
      Height = 13
      Caption = 'Mod'#232'le'
    end
    object bModele: TToolbarButton97
      Left = 586
      Top = 31
      Width = 19
      Height = 19
      Hint = 'Modifier le mod'#232'le d'#39'impression'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 0
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bModeleClick
      GlobalIndexImage = 'Z0008_S16G1'
    end
    object HLabel2: THLabel
      Left = 11
      Top = 34
      Width = 68
      Height = 13
      Caption = 'Libell'#233' compte'
      FocusControl = CompteBanque
    end
    object HCompte: THLabel
      Left = 107
      Top = 34
      Width = 208
      Height = 13
      AutoSize = False
    end
    object HModele: TPanel
      Left = 371
      Top = 32
      Width = 212
      Height = 18
      Alignment = taLeftJustify
      BevelOuter = bvLowered
      TabOrder = 1
    end
    object E_MODELE: THValComboBox
      Left = 371
      Top = 30
      Width = 212
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 3
      OnChange = E_MODELEChange
      TagDispatch = 0
      Plus = 'CLT'
      DataType = 'CPMODELESENCADECA'
    end
    object HBanque: TPanel
      Left = 371
      Top = 5
      Width = 212
      Height = 18
      Alignment = taLeftJustify
      BevelOuter = bvLowered
      TabOrder = 0
    end
    object CompteBanque: THCritMaskEdit
      Left = 107
      Top = 4
      Width = 212
      Height = 21
      TabOrder = 2
      OnDblClick = CompteBanqueDblClick
      OnExit = CompteBanqueExit
      TagDispatch = 0
      DataType = 'TZGBANQUE'
      ElipsisButton = True
      ElipsisAutoHide = True
      Libelle = HCompte
    end
  end
  object Panel1: TPanel
    Left = 294
    Top = 79
    Width = 34
    Height = 269
    Align = alRight
    BevelOuter = bvNone
    TabOrder = 4
    object bDown: TToolbarButton97
      Left = 4
      Top = 74
      Width = 27
      Height = 27
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 5
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bDownClick
      GlobalIndexImage = 'Z1206_S16G1'
    end
    object bUp: TToolbarButton97
      Left = 4
      Top = 105
      Width = 27
      Height = 27
      Hint = 'Supprimer le ch'#232'que s'#233'lectionn'#233
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 5
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bUpClick
      GlobalIndexImage = 'Z0818_S16G1'
    end
    object bGroupTiers: TToolbarButton97
      Left = 4
      Top = 140
      Width = 27
      Height = 27
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 5
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bGroupTiersClick
      GlobalIndexImage = 'Z0192_S16G1'
    end
    object bCopyDetail: TToolbarButton97
      Left = 4
      Top = 171
      Width = 27
      Height = 27
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 5
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bCopyDetailClick
      GlobalIndexImage = 'Z0190_S16G2'
    end
    object bGroupTiersEche: TToolbarButton97
      Left = 3
      Top = 203
      Width = 27
      Height = 27
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 5
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Visible = False
      OnClick = bGroupTiersEcheClick
      GlobalIndexImage = 'Z0190_S16G2'
    end
  end
  object Panel7: TPanel
    Left = 0
    Top = 56
    Width = 644
    Height = 23
    Align = alTop
    TabOrder = 5
    object HLabel1: THLabel
      Left = 2
      Top = 4
      Width = 259
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Caption = 'Liste des '#233'ch'#233'ances d'#39'origine'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HTitreImprime: THLabel
      Left = 299
      Top = 4
      Width = 315
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 463
    Top = 216
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      #233'ch'#233'ance(s) s'#233'lectionn'#233'e(s)'
      
        '1;?caption?;Certains '#233'l'#233'ments ont '#233't'#233' imprim'#233's, confirmez-vous l' +
        #39'annulation ?;Q;YN;Y;N;'
      'ch'#232'que(s) s'#233'lectionn'#233'(s)'
      '3;?caption?;Les tiers doivent '#234'tre identiques.;W;O;O;O;'
      '4;?caption?;Vous n'#39'avez rien imprim'#233'.;W;O;O;O;'
      '5;?caption?;Le montant ne peut '#234'tre n'#233'gatif ou nul.;W;O;O;O;'
      'Regroupement par tiers'
      'Regroupement par mouvement'
      'Liste des traites '#224' imprimer'
      '9;?caption?;Vous devez choisir un mod'#232'le de traite.;W;O;O;O;'
      '10;?caption?;Vous devez choisir un compte de banque.;W;O;O;O;'
      'Imprimer les traites'
      'Impression des lettres-traite'
      'Liste des BOR '#224' imprimer'
      'Imprimer les BOR'
      'Impression des lettres-BOR'
      'Transformer l'#39#233'ch'#233'ance s'#233'lectionn'#233'e en BOR'
      'Supprimer le BOR s'#233'lectionn'#233
      'G'#233'n'#233'rer un BOR par tiers et banque pr'#233'visionnelle'
      'G'#233'n'#233'rer un BOR par mouvement'
      
        '20;?caption?;La devise du journal de banque diff'#232're de celle des' +
        ' '#233'ch'#233'ances.;W;O;O;O;'
      'Regroupement par tiers et date d'#39#233'ch'#233'ance'
      'G'#233'n'#233'rer un BOR par tiers et date d'#39#233'ch'#233'ance'
      'Liste des virements '#224' imprimer'
      'Liste des pr'#233'l'#232'vements '#224' imprimer'
      'Imprimer les virements'
      'Imprimer les pr'#233'l'#232'vements'
      'Impression des virements'
      'Impression des pr'#233'l'#232'vements'
      'Liste des ch'#232'ques '#224' imprimer'
      'Transformer l'#39#233'ch'#233'ance s'#233'lectionn'#233'e en traite'
      'Transformer l'#39#233'ch'#233'ance s'#233'lectionn'#233'e en ch'#232'que'
      'Transformer l'#39#233'ch'#233'ance s'#233'lectionn'#233'e en virement'
      'Transformer l'#39#233'ch'#233'ance s'#233'lectionn'#233'e en pr'#233'l'#233'vement'
      'Supprimer la traite s'#233'lectionn'#233'e'
      'Supprimer le ch'#232'que s'#233'lectionn'#233
      'Supprimer le virement s'#233'lectionn'#233
      'Supprimer le pr'#233'l'#233'vement s'#233'lectionn'#233
      'G'#233'n'#233'rer une traite par tiers et banque pr'#233'visionnelle'
      'G'#233'n'#233'rer un ch'#232'que par tiers'
      'G'#233'n'#233'rer un virement par tiers'
      'G'#233'n'#233'rer un pr'#233'l'#233'vement par tiers'
      'G'#233'n'#233'rer une traite par mouvement'
      'G'#233'n'#233'rer un ch'#232'que par mouvement'
      'G'#233'n'#233'rer un virement par mouvement'
      'G'#233'n'#233'rer un pr'#233'l'#233'vement par mouvement'
      'G'#233'n'#233'rer une traite par tiers et date d'#39#233'ch'#233'ance'
      'G'#233'n'#233'rer un ch'#232'que par tiers et date d'#39#233'ch'#233'ance'
      'G'#233'n'#233'rer un virement par tiers et date d'#39#233'ch'#233'ance'
      'G'#233'n'#233'rer un pr'#233'l'#233'vement par tiers et date d'#39#233'ch'#233'ance'
      'Imprimer les ch'#232'ques'
      'Impression des lettres-ch'#232'que')
    Left = 419
    Top = 216
  end
end

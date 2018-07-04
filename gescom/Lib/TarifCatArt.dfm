object FTarifCatArt: TFTarifCatArt
  Left = 75
  Top = 140
  Width = 629
  Height = 441
  HelpContext = 110000059
  Caption = 'Tarifs Cat'#233'gorie article'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PENTETE: THPanel
    Left = 0
    Top = 0
    Width = 613
    Height = 89
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_TARIFARTICLE: THLabel
      Left = 12
      Top = 12
      Width = 76
      Height = 13
      Caption = 'Cat'#233'gorie &article'
      FocusControl = GF_TARIFARTICLE
    end
    object TGF_DEVISE: THLabel
      Left = 12
      Top = 39
      Width = 59
      Height = 13
      Caption = 'Code &devise'
      FocusControl = GF_DEVISE
    end
    object GF_TARIFARTICLE: THValComboBox
      Left = 97
      Top = 8
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = GF_TARIFARTICLEChange
      TagDispatch = 0
      DataType = 'GCTARIFARTICLE'
    end
    object GF_DEVISE: THValComboBox
      Left = 97
      Top = 35
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = GF_DEVISEChange
      TagDispatch = 0
      DataType = 'TTDEVISE'
    end
    object PTITRE: THPanel
      Left = 0
      Top = 64
      Width = 621
      Height = 25
      Caption = 'PTITRE'
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 2
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object TTYPETARIF: THLabel
        Left = 8
        Top = 6
        Width = 66
        Height = 13
        Caption = 'TTYPETARIF'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
    end
    object CBCATCLI: TCheckBox
      Left = 272
      Top = 37
      Width = 157
      Height = 17
      Caption = '&Tarif par cat'#233'gorie client'
      TabOrder = 3
      OnClick = CBCATCLIClick
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 367
    Width = 613
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object Toolbar972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 613
      ClientAreaHeight = 32
      ClientAreaWidth = 613
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Tag = 1
        Left = 569
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
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
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 538
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 506
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer les tarifs'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
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
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 14
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher '
        Caption = 'Rechercher'
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
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BCollerCond: TToolbarButton97
        Left = 160
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Coller conditions'
        Caption = 'Coller'
        DisplayMode = dmGlyphOnly
        Flat = False
        OnClick = BCollerCondClick
        GlobalIndexImage = 'Z0175_S16G1'
      end
      object BCopierCond: TToolbarButton97
        Left = 129
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Copier conditions'
        Caption = 'Copier'
        DisplayMode = dmGlyphOnly
        Flat = False
        OnClick = BCopierCondClick
        GlobalIndexImage = 'Z0342_S16G1'
      end
      object BCondAplli: TToolbarButton97
        Left = 98
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Modifier conditions'
        Caption = 'Modifier'
        DisplayMode = dmGlyphOnly
        Flat = False
        NumGlyphs = 2
        OnClick = BCondAplliClick
        GlobalIndexImage = 'Z0105_S16G2'
      end
      object BVoirCond: TToolbarButton97
        Tag = 1
        Left = 67
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Voir conditions tarifaires'
        Caption = 'Voir condition'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Spacing = -1
        OnClick = BVoirCondClick
        GlobalIndexImage = 'Z0206_S16G1'
        IsControl = True
      end
    end
  end
  object PPIED: THPanel
    Left = 0
    Top = 313
    Width = 613
    Height = 54
    Align = alBottom
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_CASCADEREMISE: THLabel
      Left = 11
      Top = 6
      Width = 60
      Height = 13
      Caption = '&Mode remise'
      FocusControl = GF_CASCADEREMISE
    end
    object TGF_REMISE: THLabel
      Left = 11
      Top = 29
      Width = 84
      Height = 13
      Caption = 'Remise r'#233'sultante'
      FocusControl = GF_REMISE
    end
    object GF_CASCADEREMISE: THValComboBox
      Left = 105
      Top = 3
      Width = 126
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = GF_CASCADEREMISEChange
      TagDispatch = 0
      DataType = 'GCCASCADEREMISE'
    end
    object GF_REMISE: THCritMaskEdit
      Left = 105
      Top = 26
      Width = 126
      Height = 21
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
      OnChange = GF_REMISEChange
      TagDispatch = 0
      OpeType = otReel
    end
    object CONDAPPLIC: THRichEditOLE
      Left = 528
      Top = 8
      Width = 73
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      Visible = False
      WordWrap = False
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
          'ial;}}'
        '{\colortbl ;\red255\green0\blue0;}'
        
          '{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\cf1\f0\' +
          'fs16 CONDAPPLIC'
        '\par }')
    end
  end
  object PTARIF: THPanel
    Left = 0
    Top = 89
    Width = 613
    Height = 224
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_CatA: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 611
      Height = 222
      Align = alClient
      ColCount = 8
      DefaultRowHeight = 18
      RowCount = 3
      TabOrder = 0
      OnEnter = G_CatAEnter
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_CatARowEnter
      OnRowExit = G_CatARowExit
      OnCellEnter = G_CatACellEnter
      OnCellExit = G_CatACellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = G_CatAElipsisClick
    end
  end
  object PCATEGORIE: THPanel
    Left = 0
    Top = 89
    Width = 613
    Height = 224
    Align = alClient
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_CatC: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 611
      Height = 222
      Align = alClient
      ColCount = 8
      DefaultRowHeight = 18
      RowCount = 3
      TabOrder = 0
      OnEnter = G_CatCEnter
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_CatCRowEnter
      OnRowExit = G_CatCRowExit
      OnCellEnter = G_CatCCellEnter
      OnCellExit = G_CatCCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = G_CatCElipsisClick
      RowHeights = (
        18
        18
        18)
    end
  end
  object TCONDTARF: TToolWindow97
    Left = 88
    Top = 116
    ClientHeight = 175
    ClientWidth = 516
    Caption = 'Conditions tarifaires'
    ClientAreaHeight = 175
    ClientAreaWidth = 516
    Resizable = False
    TabOrder = 5
    Visible = False
    OnClose = TCONDTARFClose
    object G_COND: THGrid
      Left = 0
      Top = 0
      Width = 516
      Height = 175
      Align = alClient
      ColCount = 4
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 25
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Table'
        'Champ'
        'Condition'
        'Valeur')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
      ColWidths = (
        73
        162
        104
        154)
    end
    object FComboTIE: THValComboBox
      Left = 207
      Top = 21
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 1
      Visible = False
      TagDispatch = 0
    end
    object FComboART: THValComboBox
      Left = 207
      Top = 46
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 2
      Visible = False
      TagDispatch = 0
    end
    object FComboLIG: THValComboBox
      Left = 207
      Top = 72
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 3
      Visible = False
      TagDispatch = 0
    end
    object FComboPIE: THValComboBox
      Left = 207
      Top = 97
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 4
      Visible = False
      TagDispatch = 0
    end
    object GF_CONDAPPLIC: THRichEditOLE
      Left = 272
      Top = 56
      Width = 185
      Height = 89
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      Visible = False
      WordWrap = False
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
          'ial;}}'
        
          '{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\f0\fs16' +
          ' GF_CONDAPPLIC'
        '\par }')
    end
    object FTable: THValComboBox
      Left = 207
      Top = 121
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 6
      Visible = False
      Items.Strings = (
        'Piece'
        'Ligne'
        'Tiers'
        'Article')
      TagDispatch = 0
      Values.Strings = (
        'PIECE'
        'LIGNE'
        'TIERS'
        'ARTICLE')
    end
    object FOpe: THValComboBox
      Left = 207
      Top = 146
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Ctl3D = False
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 7
      Visible = False
      TagDispatch = 0
      DataType = 'TTCOMPARE'
    end
  end
  object POPZ: TPopupMenu
    Left = 436
    Top = 202
    object InfArticle: TMenuItem
      Caption = 'Article'
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      '1;?caption?;Voulez-vous supprimer l'#39'enregistrement ?;Q;YNC;N;C;'
      '2;?caption?;Les bornes de quantit'#233's sont incoh'#233'rentes.;W;O;O;O;'
      
        '3;?caption?;Les bornes de dates de validit'#233' sont incoh'#233'rentes.;W' +
        ';O;O;O;'
      '4;?caption?;Ce code d'#233'p'#244't n'#39'existe pas.;W;O;O;O;'
      '5;?caption?;Cette cat'#233'gorie de tiers n'#39'existe pas.;W;O;O;O;'
      '6;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;Y;N;')
    Left = 471
    Top = 202
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 516
    Top = 202
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 563
    Top = 202
  end
  object HMess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ATTENTION : Tarif non enregistr'#233' !'
      
        'ATTENTION : Ce tarif, en cours de traitement par un autre utilis' +
        'ateur, n'#39'a pas '#233't'#233' enregistr'#233'e !')
    Left = 396
    Top = 202
  end
  object HTitre: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Cat'#233'gorie article'
      'Cat'#233'gorie de client'
      'Aucun type de tarif s'#233'lectionn'#233
      'Tarif global')
    Left = 355
    Top = 202
  end
end

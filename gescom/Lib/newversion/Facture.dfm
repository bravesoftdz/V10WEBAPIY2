object FFacture: TFFacture
  Left = 287
  Top = 53
  Width = 1011
  Height = 698
  HelpContext = 119000017
  ActiveControl = GP_DATEPIECE
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'traitance'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  OnCanResize = FormCanResize
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object BOuvrirFermer: TToolbarButton97
    Left = 168
    Top = 288
    Width = 20
    Height = 20
    DisplayMode = dmGlyphOnly
    Images = ImgListPlusMoins
    Opaque = False
    OnClick = BOuvrirFermerClick
  end
  object PEntete: THPanel
    Left = 0
    Top = 0
    Width = 995
    Height = 125
    Align = alTop
    FullRepaint = False
    TabOrder = 0
    OnExit = PEnteteExit
    BackGroundEffect = bdFond
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    ColorNb = 2
    TextEffect = tenone
    AutoResize = True
    object BDevise: TToolbarButton97
      Tag = -100
      Left = 226
      Top = 58
      Width = 61
      Height = 21
      Caption = 'Devise'
      Alignment = taLeftJustify
      DisplayMode = dmTextOnly
      DropdownArrow = True
      DropdownMenu = PopD
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Layout = blGlyphTop
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      IsControl = True
    end
    object HMOTIFMVT: THLabel
      Left = 236
      Top = 62
      Width = 23
      Height = 13
      Caption = 'Motif'
      FocusControl = MOTIFMVT
      Transparent = True
      Visible = False
    end
    object HGP_NUMPIECE: THLabel
      Left = 8
      Top = 37
      Width = 12
      Height = 13
      Caption = 'N'#176
      Transparent = True
    end
    object HGP_REFINTERNE: THLabel
      Left = 491
      Top = 10
      Width = 50
      Height = 13
      Caption = 'R'#233'f'#233'rence'
      FocusControl = GP_REFINTERNE
      Transparent = True
    end
    object HGP_TIERS: THLabel
      Left = 236
      Top = 10
      Width = 23
      Height = 13
      Caption = 'Tiers'
      FocusControl = GP_TIERS
      Transparent = True
    end
    object HGP_DATEPIECE: THLabel
      Left = 120
      Top = 37
      Width = 12
      Height = 13
      Caption = 'du'
      FocusControl = GP_DATEPIECE
      Transparent = True
    end
    object HGP_REPRESENTANT: THLabel
      Left = 491
      Top = 88
      Width = 54
      Height = 13
      Caption = 'Commercial'
      FocusControl = GP_REPRESENTANT
      Transparent = True
    end
    object HGP_ETABLISSEMENT: THLabel
      Left = 236
      Top = 36
      Width = 65
      Height = 13
      Caption = 'Etablissement'
      FocusControl = GP_ETABLISSEMENT
      Transparent = True
    end
    object FTitrePiece: THLabel
      Left = 5
      Top = 6
      Width = 226
      Height = 20
      Alignment = taCenter
      AutoSize = False
      Caption = 'PIECE COMMERCIALE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
      OnClick = FTitrePieceClick
      Ombre = True
      OmbreDecalX = 1
      OmbreDecalY = 1
      OmbreColor = clGray
    end
    object HGP_DEPOT: THLabel
      Left = 491
      Top = 36
      Width = 29
      Height = 13
      Caption = 'D'#233'p'#244't'
      FocusControl = GP_DEPOT
      Transparent = True
    end
    object ISigneEuro: TImage
      Left = 291
      Top = 60
      Width = 16
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        0000100000000100040000000000800000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
        FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
        4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
        FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
        4FF4}
      Stretch = True
      Transparent = True
      Visible = False
    end
    object HGP_DATELIVRAISON: THLabel
      Left = 491
      Top = 62
      Width = 53
      Height = 13
      Caption = 'Livraison le'
      FocusControl = GP_DATELIVRAISON
      Transparent = True
    end
    object HGP_AFFAIRE: THLabel
      Left = 236
      Top = 88
      Width = 30
      Height = 13
      Caption = 'Affaire'
      Transparent = True
    end
    object BRechAffaire: TToolbarButton97
      Left = 288
      Top = 84
      Width = 19
      Height = 22
      Hint = 'Rechercher client/affaire'
      Opaque = False
      OnClick = BRechAffaire__Click
      GlobalIndexImage = 'Z0002_S16G1'
    end
    object HGP_DOMAINE: THLabel
      Left = 12
      Top = 88
      Width = 63
      Height = 13
      Caption = 'Domaine act.'
      FocusControl = GP_DOMAINE
      Transparent = True
      Visible = False
    end
    object HGP_REFEXTERNE: THLabel
      Left = 491
      Top = 10
      Width = 50
      Height = 13
      Caption = 'R'#233'f'#233'rence'
      FocusControl = GP_REFINTERNE
      Transparent = True
    end
    object GP_REFINTERNE: TEdit
      Left = 552
      Top = 6
      Width = 136
      Height = 21
      MaxLength = 35
      TabOrder = 22
      OnExit = GP_REFINTERNEExit
    end
    object GP_DATEREFEXTERNE: THCritMaskEdit
      Left = 624
      Top = 6
      Width = 64
      Height = 21
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 8
      Text = '  /  /    '
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object GP_REFEXTERNE: TEdit
      Left = 552
      Top = 6
      Width = 66
      Height = 21
      MaxLength = 35
      TabOrder = 7
      OnExit = GP_REFEXTERNEExit
    end
    object MOTIFMVT: THValComboBox
      Left = 310
      Top = 58
      Width = 163
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 26
      Visible = False
      OnChange = GP_ETABLISSEMENTChange
      TagDispatch = 0
      DataType = 'GCMOTIFMOUVEMENT'
    end
    object GP_DEVISE: THValComboBox
      Left = 310
      Top = 58
      Width = 163
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 21
      OnChange = GP_DEVISEChange
      TagDispatch = 0
      DataType = 'TTDEVISE'
    end
    object EUROPIVOT: TEdit
      Left = 310
      Top = 57
      Width = 162
      Height = 21
      Enabled = False
      TabOrder = 25
      Visible = False
    end
    object LDevise: THPanel
      Left = 291
      Top = 60
      Width = 17
      Height = 16
      BevelOuter = bvNone
      Caption = '$'
      Ctl3D = True
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 24
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object GP_TIERS: THCritMaskEdit
      Left = 310
      Top = 8
      Width = 163
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
      OnDblClick = GP_TIERSDblClick
      OnEnter = GP_TIERSEnter
      OnExit = GP_TIERSExit
      TagDispatch = 0
      DataType = 'GCTIERSSAISIE'
      ElipsisButton = True
      OnElipsisClick = GP_TIERSElipsisClick
    end
    object GP_DATEPIECE: THCritMaskEdit
      Left = 139
      Top = 32
      Width = 78
      Height = 21
      AutoSelect = False
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 0
      Text = '  /  /    '
      OnExit = GP_DATEPIECEExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object GP_REPRESENTANT: THCritMaskEdit
      Left = 551
      Top = 84
      Width = 137
      Height = 21
      Enabled = False
      TabOrder = 11
      OnEnter = GP_REPRESENTANTEnter
      OnExit = GP_REPRESENTANTExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GP_REPRESENTANTElipsisClick
    end
    object GP_NUMEROPIECE: THPanel
      Left = 35
      Top = 33
      Width = 80
      Height = 21
      BevelOuter = bvLowered
      Caption = 'Non affect'#233
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 23
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object LIBELLETIERS: THPanel
      Left = 11
      Top = 105
      Width = 155
      Height = 16
      Alignment = taLeftJustify
      BevelOuter = bvNone
      Ctl3D = True
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 20
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object GP_ETABLISSEMENT: THValComboBox
      Left = 310
      Top = 32
      Width = 163
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      OnChange = GP_ETABLISSEMENTChange
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
    end
    object GP_DEPOT: THValComboBox
      Left = 551
      Top = 32
      Width = 137
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 9
      OnChange = GP_DEPOTChange
      TagDispatch = 0
      DataType = 'GCDEPOT'
    end
    object GP_SAISIECONTRE: TCheckBox
      Left = 84
      Top = 60
      Width = 32
      Height = 17
      TabStop = False
      Caption = 'GP_SAISIECONTRE'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 16
      Visible = False
    end
    object GP_FACTUREHT: TCheckBox
      Left = 124
      Top = 60
      Width = 28
      Height = 17
      TabStop = False
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
      TabOrder = 17
      Visible = False
    end
    object GP_REGIMETAXE: THValComboBox
      Left = 157
      Top = 58
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
      TabOrder = 18
      TabStop = False
      Visible = False
      OnChange = GP_REGIMETAXEChange
      TagDispatch = 0
      DataType = 'TTREGIMETVA'
    end
    object GP_DATELIVRAISON: THCritMaskEdit
      Left = 551
      Top = 58
      Width = 137
      Height = 21
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 10
      Text = '  /  /    '
      OnExit = GP_DATELIVRAISONExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object GP_AFFAIRE0: THCritMaskEdit
      Left = 54
      Top = 57
      Width = 27
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 15
      Visible = False
      TagDispatch = 0
    end
    object GP_AFFAIRE1: THCritMaskEdit
      Left = 310
      Top = 84
      Width = 44
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 3
      OnChange = GP_AFFAIRE1__Change
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AFFAIRE2: THCritMaskEdit
      Left = 358
      Top = 84
      Width = 49
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 4
      OnChange = GP_AFFAIRE2__Change
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AFFAIRE3: THCritMaskEdit
      Left = 410
      Top = 84
      Width = 41
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 5
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AVENANT: THCritMaskEdit
      Left = 455
      Top = 84
      Width = 18
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 6
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object AFF_TIERS: THCritMaskEdit
      Left = 36
      Top = 65
      Width = 16
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 14
      Visible = False
      OnChange = GP_AFFAIRE1__Change
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_TVAENCAISSEMENT: THValComboBox
      Left = 190
      Top = 58
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
      TabOrder = 19
      TabStop = False
      Visible = False
      OnChange = GP_TVAENCAISSEMENTChange
      TagDispatch = 0
      DataType = 'GCTVAENCAISSEMENT'
    end
    object GP_DOMAINE: THValComboBox
      Left = 90
      Top = 84
      Width = 127
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 12
      Visible = False
      OnChange = GP_DOMAINEChange
      TagDispatch = 0
      Vide = True
      DataType = 'YYDOMAINE'
    end
    object GP_HRDOSSIER: THCritMaskEdit
      Left = 12
      Top = 64
      Width = 17
      Height = 21
      TabStop = False
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 13
      Visible = False
      TagDispatch = 0
    end
  end
  object TSaisieAveugle: TToolWindow97
    Left = 18
    Top = 148
    ClientHeight = 117
    ClientWidth = 657
    Caption = 'Saisie par codes '#224' barres'
    ClientAreaHeight = 117
    ClientAreaWidth = 657
    DockableTo = []
    DockPos = 0
    DragHandleStyle = dhNone
    FullSize = True
    TabOrder = 13
    Visible = False
    OnClose = TSaisieAveugleClose
    OnResize = TSaisieAveugleResize
    object GSAveugle: THGrid
      Left = 0
      Top = 0
      Width = 657
      Height = 68
      Align = alClient
      ColCount = 4
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs]
      TabOrder = 0
      OnKeyDown = GSAveugleKeyDown
      SortedCol = -1
      Titres.Strings = (
        ''
        'Code barre'
        'D'#233'signation'
        'Quantit'#233)
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GSAveugleRowEnter
      OnRowExit = GSAveugleRowExit
      OnCellEnter = GSAveugleCellEnter
      OnCellExit = GSAveugleCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      RowHeights = (
        18
        18)
    end
    object PSaisieCB: THPanel
      Left = 0
      Top = 86
      Width = 657
      Height = 31
      Align = alBottom
      BevelOuter = bvNone
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      DesignSize = (
        657
        31)
      object BtValiderSA: TToolbarButton97
        Left = 594
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider la saisie par code barre'
        Anchors = [akTop, akRight, akBottom]
        DisplayMode = dmGlyphOnly
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          1800000000000003000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828484000084000000828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082848400000082000082008400000082840082840082840082840082840082
          8400828400828400828400828400828484000000820000820000820000820084
          0000008284008284008284008284008284008284008284008284008284840000
          0082000082000082000082000082000082008400000082840082840082840082
          8400828400828400828484000000820000820000820000FF0000820000820000
          8200008200840000008284008284008284008284008284008284008200008200
          00820000FF0000828400FF000082000082000082008400000082840082840082
          8400828400828400828400FF0000820000FF0000828400828400828400FF0000
          820000820000820084000000828400828400828400828400828400828400FF00
          00828400828400828400828400828400FF000082000082000082008400000082
          8400828400828400828400828400828400828400828400828400828400828400
          828400FF00008200008200008200840000008284008284008284008284008284
          00828400828400828400828400828400828400828400FF000082000082000082
          0084000000828400828400828400828400828400828400828400828400828400
          828400828400828400FF00008200008200008200840000008284008284008284
          00828400828400828400828400828400828400828400828400828400FF000082
          0000820000820084000000828400828400828400828400828400828400828400
          828400828400828400828400828400FF00008200008200840000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400FF0000820000820000828400828400828400828400828400828400828400
          828400828400828400828400828400828400828400FF00008284}
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        Opaque = False
        ParentColor = True
        ParentShowHint = False
        ShowHint = True
        OnClick = BtValiderSAClick
      end
      object BtFermerSA: TToolbarButton97
        Left = 622
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight, akBottom]
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentColor = True
        ParentShowHint = False
        ShowHint = True
        OnClick = BtFermerSAClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BtRefreshSA: TToolbarButton97
        Left = 0
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Vider grille'
        Anchors = [akLeft, akTop, akBottom]
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentColor = True
        ParentShowHint = False
        ShowHint = True
        OnClick = BtRefreshSAClick
        GlobalIndexImage = 'Z0050_S16G1'
      end
      object BtDelLigneSA: TToolbarButton97
        Left = 28
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer ligne'
        Anchors = [akLeft, akTop, akBottom]
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentColor = True
        ParentShowHint = False
        ShowHint = True
        OnClick = BtDelLigneSAClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BtRetour: TToolbarButton97
        Left = 56
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Retour article'
        AllowAllUp = True
        Anchors = [akLeft, akTop, akBottom]
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentColor = True
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0192_S16G1'
      end
      object BTypeArticle: TToolbarButton97
        Left = 188
        Top = 0
        Width = 19
        Height = 22
        NoBorder = True
        Visible = False
        OnClick = BTypeArticleClick
      end
      object BtTPI: TToolbarButton97
        Left = 84
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'vidage TPI'
        Anchors = [akLeft, akTop, akBottom]
        DisplayMode = dmGlyphOnly
        Glyph.Data = {
          76050000424D7605000000000000360000002800000015000000150000000100
          180000000000400500000000000000000000000000000000000000FFFF000000
          00000000FFFF00FFFF00FFFF00FFFF80808080808000FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF000000
          00000000FFFF00FFFF80808000000000000000000080808000FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF000000
          00000000000080808000000080808080808080808080808000FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF000000
          000000000000808080C0C0C0C0C0C0C0C0C0C0C0C080808080808080808000FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          808080808080FFFFFFC0C0C0C0C0C0C0C0C0FFFFFFFFFFFF00000080808000FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF000000
          808080FFFFFFC0C0C0FFFFFFC0C0C0C0C0C0C0C0C0C0C0C0C0C0C08080808080
          8080808080808080808080808080808000FFFF00FFFF00FFFF0000FFFF000000
          808080FFFFFFFFFFFFC0C0C0FFFFFFC0C0C0C0C0C0C0C0C0C0C0C0C0C0C00000
          0080808000000000000000000000000080808000FFFF00FFFF00000000808080
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC0C0C0FFFFFF808080FFFFFFFFFFFFFFFF
          FF808080808080808080FFFFFFFFFFFF00000080808000FFFF0000FFFF000000
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF808080000000FFFFFFC0C0C0FFFF
          FFFFFFFF0000000000000000000000008080800000008080800000FFFF00FFFF
          000000FFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFF0000000000000000
          00000000000000000000000000808000000000FFFFFF8080800000FFFF00FFFF
          00FFFFC0C0C0C0C0C0FFFFFFFFFFFF0000000000000000000000000000000000
          0000000000000000000080800080800000000000000000FFFF0000FFFF00FFFF
          00FFFF00FFFF000000C0C0C0FFFFFFFFFFFFFFFFFF0000000000000000000000
          0000000080800080800080800080800000000000FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF000000FFFFFFFFFFFF0000000000000000000000000000
          00808000FFFF0080800000000000000000FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF000000C0C0C00000000000000000000000008080
          0080800080800080800000000000FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF000000808080000000000000000000000000808000FFFF
          0080800000000000000000FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF000000FFFFFF000000FFFFFF0000008080008080008080
          0080800000000000FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF000000FFFFFF000000FFFFFF808000FFFFFF8080000000
          0000000000FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF000000FFFFFF0000000000008080008080008080000000
          0000FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF00000000FFFFFFFFFFFFFFFF00000000000000000000FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF000000000000FFFFFF00000000000000FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00}
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        Opaque = False
        ParentColor = True
        ParentShowHint = False
        ShowHint = True
        OnClick = BtTPIClick
      end
    end
    object HTotaux: THPanel
      Left = 0
      Top = 68
      Width = 657
      Height = 18
      Align = alBottom
      Alignment = taLeftJustify
      BevelInner = bvLowered
      BevelOuter = bvLowered
      Caption = 'Totaux'
      Color = clWindow
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      TabOrder = 2
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object SommeQTE: THNumEdit
        Left = 296
        Top = 0
        Width = 57
        Height = 21
        ReadOnly = True
        TabOrder = 0
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
  end
  object BZoomArticle: THBitBtn
    Tag = 100
    Left = 27
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Voir article'
    Caption = 'Article'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    Visible = False
    OnClick = BZoomArticleClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomTiers: THBitBtn
    Tag = 100
    Left = 58
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Voir tiers'
    Caption = 'Tiers'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    Visible = False
    OnClick = BZoomTiersClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomTarif: THBitBtn
    Tag = 100
    Left = 90
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Explication du tarif'
    Caption = 'Tarif'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = BZoomTarifClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomCommercial: THBitBtn
    Tag = 100
    Left = 150
    Top = 148
    Width = 27
    Height = 29
    Hint = 'Voir commercial'
    Caption = 'Commercial'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Visible = False
    OnClick = BZoomCommercialClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomOrigine: THBitBtn
    Tag = 100
    Left = 194
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Pi'#232'ce d'#39'origine'
    Caption = 'Origine'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
    Visible = False
    OnClick = BZoomOrigineClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomSuivante: THBitBtn
    Tag = 100
    Left = 298
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Pi'#232'ce suivante'
    Caption = 'Suivante'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 10
    Visible = False
    OnClick = BZoomSuivanteClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomEcriture: THBitBtn
    Tag = 100
    Left = 263
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Ecriture comptable'
    Caption = 'Ecriture'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 11
    Visible = False
    OnClick = BZoomEcritureClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomDevise: THBitBtn
    Tag = 100
    Left = 334
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Voir devise'
    Caption = 'Devise'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 12
    Visible = False
    OnClick = BZoomDeviseClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomStock: THBitBtn
    Tag = 100
    Left = 279
    Top = 155
    Width = 28
    Height = 27
    Hint = 'Ecriture stock'
    Caption = 'Compta. Stock'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 15
    Visible = False
    OnClick = BZoomStockClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomDispo: THBitBtn
    Tag = 100
    Left = 339
    Top = 155
    Width = 28
    Height = 27
    Hint = 'Voir disponiblit'#233
    Caption = 'Disponibilit'#233
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 16
    Visible = False
    OnClick = BZoomDispoClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomPiece: THBitBtn
    Tag = 100
    Left = 366
    Top = 155
    Width = 28
    Height = 27
    Hint = 'Synth'#232'se des pi'#232'ces'
    Caption = 'Pi'#232'ce'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 18
    Visible = False
    OnClick = BZoomPieceClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomTache: THBitBtn
    Tag = 100
    Left = 435
    Top = 155
    Width = 28
    Height = 27
    Hint = 'Voir t'#226'che de l'#39'affaire'
    Caption = 'Tache'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 19
    Visible = False
    OnClick = BZoomTacheClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomRevision: THBitBtn
    Tag = 100
    Left = 467
    Top = 155
    Width = 28
    Height = 27
    Hint = 'Voir r'#233'vision de prix'
    Caption = 'R'#233'vision prix'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 21
    Visible = False
    OnClick = BZoomRevisionClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomProposition: THBitBtn
    Tag = 100
    Left = 219
    Top = 83
    Width = 28
    Height = 27
    Hint = 'Voir Proposition'
    Caption = 'Compta. Stock'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 20
    Visible = False
    OnClick = BZoomPropositionClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object PPiedTot: TPanel
    Left = 0
    Top = 513
    Width = 995
    Height = 24
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 23
    object PGTG: TPanel
      Left = 0
      Top = -56
      Width = 995
      Height = 80
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 0
      object GTG: THGrid
        Left = 0
        Top = 0
        Width = 995
        Height = 80
        Align = alClient
        ColCount = 9
        DefaultRowHeight = 18
        Enabled = False
        FixedColor = clActiveCaption
        RowCount = 4
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
        ScrollBars = ssNone
        TabOrder = 0
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
      end
    end
    object PGT: TPanel
      Left = 0
      Top = 0
      Width = 995
      Height = 21
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      object GT: THGrid
        Left = 0
        Top = 0
        Width = 995
        Height = 22
        Align = alTop
        DefaultRowHeight = 18
        Enabled = False
        RowCount = 1
        FixedRows = 0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Options = [goEditing]
        ParentFont = False
        ScrollBars = ssNone
        TabOrder = 0
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
      end
    end
  end
  object WAIT: TToolWindow97
    Left = 232
    Top = 208
    ClientHeight = 171
    ClientWidth = 445
    Caption = 'Message'
    ClientAreaHeight = 171
    ClientAreaWidth = 445
    DragHandleStyle = dhSingle
    Resizable = False
    TabOrder = 24
    Visible = False
    object Label1: TLabel
      Left = 0
      Top = 0
      Width = 445
      Height = 24
      Align = alTop
      Alignment = taCenter
      AutoSize = False
      Caption = 'XXX'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 0
      Top = 144
      Width = 445
      Height = 27
      Align = alBottom
      Alignment = taCenter
      AutoSize = False
      Caption = 'Veuillez patienter...'
      Font.Charset = ANSI_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Animate1: TAnimate
      Left = 91
      Top = 56
      Width = 272
      Height = 60
      CommonAVI = aviCopyFile
      StopFrame = 20
      Timers = True
    end
  end
  object PGlobPied: THPanel
    Left = 0
    Top = 537
    Width = 995
    Height = 122
    Align = alBottom
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 26
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    AutoResize = True
    object PPied: THPanel
      Left = 0
      Top = 0
      Width = 995
      Height = 87
      Align = alTop
      BevelOuter = bvNone
      FullRepaint = False
      TabOrder = 1
      OnResize = PPiedResize
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      AutoResize = True
      DesignSize = (
        995
        87)
      object HGP_ESCOMPTE: THLabel
        Left = 222
        Top = 6
        Width = 58
        Height = 13
        Caption = '&Escompte %'
        Color = clInfoBk
        FocusControl = GP_ESCOMPTE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Transparent = True
      end
      object HGP_MODEREGLE: TFlashingLabel
        Left = 222
        Top = 48
        Width = 51
        Height = 13
        Caption = 'R'#232'glement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object HGP_REMISEPIED: THLabel
        Left = 222
        Top = 27
        Width = 46
        Height = 13
        Caption = '&Remise %'
        Color = clInfoBk
        FocusControl = GP_REMISEPIED
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Transparent = True
      end
      object LLIBRG: TFlashingLabel
        Left = 222
        Top = 68
        Width = 25
        Height = 13
        Caption = 'XXX'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object LMontantRg: TFlashingLabel
        Left = 222
        Top = 84
        Width = 25
        Height = 13
        Caption = 'XXX'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object BLanceCalc: TToolbarButton97
        Left = 597
        Top = 2
        Width = 405
        Height = 101
        Caption = 'Document '#224' recalculer'
        Anchors = [akTop, akRight]
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          36190000424D3619000000000000360000002800000028000000280000000100
          20000000000000190000120B0000120B00000000000000000000FFFFFF00FFFF
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
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F9FAFA00E7E8E800E3E4
          E300D0CCCC00D6D4D400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00F9FAF900ECECEC00E0E1E000D5D7D500CDD1CF00D5D9D700DDDD
          DC00E1DDDC00B3AFAE00C5C4C300FDFDFD00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FCFCFC00F2F1F100E2E1
          E100D7D7D500CECDCD00D4D4D200DDDDDB00E3E2E100DCDEDC00D7DBD800CBCA
          C900CFCAC900C2BEBD00908D8C00B1B2B100FCFCFC00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FEFEFE00F5F5F500E3E3E300DBDBDB00D0CFCE00D9D6D400DFDD
          DC00E4E1E000E5E2E100DDDAD900D5D2D100CCC9C800C3C3C200BABCBA00B9B6
          B500B1ACAB00BEB7B70096919000807E7D00BCBEBD00FAFBFB00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F8F8F800E5E5
          E500E0E0E000D0D0D000D5D5D500DCDCDC00E1E1E100E5E4E300E4E0DF00DCD8
          D700D5D1D000CDC9C800C1BDBC00BFBBBA00B3AFAE00ADADAB00A8A9A700A6A3
          A200A59E9D00B3ABAB00A6A0A0008885840097989700CBCCCC00FAFAFA00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FBFBFB00EDEDED00E4E4E400D6D6D600D8D8D800DFDF
          DF00E4E4E400E8E8E800E5E5E500DFDFDF00D6D6D600D2D0D000CDC8C700C3BE
          BD00BDB9B800B8B4B300B3AEAD00AFAAA900AFABAA00B1AEAD00B1AFAE00B6B1
          B00098909000A59C9C00B5ADAD00908D8C00A3A2A000B5B6B500D9D9D900FBFB
          FB00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FEFE
          FE00FCFCFC00E5E5E500D8D8D800DEDEDE00E6E6E600E9EAEA00EAEBEB00E6E7
          E700DEDEDE00D6D6D600D2D2D200CBCBCB00C4C4C400C0BDBD00BEB9B800BAB5
          B400B2ADAC00B5B0AF00BFBAB900C5C0BF00D4CECD00DCD8D700E7E6E500F2ED
          EC00AAA0A1009C919200B5AEAD009A959400A8A6A500C1C1C000D2D2D200E2E2
          E200FAFAFA00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FAFA
          FA00F5F5F500ECECEC00F0F0F000EEEEEE00E6E6E600E2E2E200DAD9D900D7D6
          D500D2D1D100CDCDCC00C8C7C600C4C3C300B7B6B600BBB9B900C6C2C100C5C1
          C000ABA7A600AAA6A500DAD6D500F1EDEC00D7D3D200C5C0BF00DFDBD900DDD8
          D600C1BAB9009A929100AFA8A700A7A2A000A7A4A200CDCACA00DFDCDC00E4E1
          E200F0EEEE00FDFDFD00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FEFE
          FE00FCFCFC00EBEBEB00E2E2E200DCDCDC00DADADA00D9D8D800D6D4D300D3D1
          D000C7C5C300BEBCBB00CAC8C700CECCCB00CBC9C800ABA9A800DAD8D700EFED
          EC00D3D1D0009C9A9900C8C6C500D5D3D200CCCAC900C7C3C100C8C3BF00B8B3
          B000ADA8A5009C979400ACA8A400B1ACA900AEA9A600EDE7E700F6F1F200FAF7
          F800FDF9FA00FEFDFD00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00E9E9E900DADADA00DCDCDC00D0D1D100BEBEBD00C2C0BF00CECC
          CB00D4D2D100BDBBBA00BDBBBA00E6E4E300E8E6E500B5B3B200C6C4C300D1CF
          CE00D5D3D200AEACAB00B2B0AF00B4B2B100B6B5B300A5A19F00AFAAA700BBB6
          B300CCC7C400A19C9900A09B9800B9B4B100AFACA800FBFBFA00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00F5F5F500DDDDDD00D4D4D400D1D1D100D2D1D100AEACAB00D8D6
          D500B5B3B200C4C2C100B7B5B400C9C7C600CCCAC900C4C2C100B9B7B600BFBD
          BC00B7B5B400A9A7A600B2B0AF00C2C0BF00D5D4D300C8C5C300A8A3A000E2DD
          DA00D6D1CE00BEB9B60098939000B9B4B100A19D9A00F1F0F000FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FDFDFD00E0E0E000DADADA00C9C9C900B1B1B000C2C0BF00C3C1
          C000C1BFBE00C8C6C500C1BFBE00C5C3C200BDBBBA00BBB9B800B8B6B500CAC8
          C700CBC9C800C2C0BF00A3A1A000D5D3D200DAD8D800D6D3D000A9A3A000CDC8
          C500C9C4C100D4CFCC00A5A09D00B2ADAA00AEA9A600D0CDCD00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00EAEAEA00DCDCDC00CECECE00CDCCCC00C5C3C200C8C6
          C500CBC9C800C7C5C400B9B7B600C0BEBD00D7D5D400D2D0CF00B3B1B000D2D0
          CF00D5D3D200C3C1C000ABA9A800C2C0BF00CCCAC900D2CECC00BAB5B200B7B2
          AF00BFBAB700B9B4B100A49F9C00B0ABA800B6B1AE00B0ACAC00FEFEFE00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00F8F8F800E5E5E500D9D8D800C9C8C800D0CECD00B7B5
          B400D1CFCE00DCDBDA00D2D0CF00B1AFAE00DDDBDA00B9B7B600C0BEBD00BDBB
          BA00C7C5C400C6C4C300BBB9B800BBB9B800BAB9B800B9B6B300B5B0AD00AEA9
          A600CEC9C600DAD5D200B4AFAC00A19C9900BFBBB80096909000F3F2F200FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FCFCFC00F1F1F100EDEDED00D9D9D900C6C4C300B9B7
          B600C3C1C000CECCCB00BDBBBA00C0BEBD00CAC8C700C6C4C400CAC8C700B7B5
          B400C5C3C300C0BEBD00BAB8B700ACAAA900CFCDCC00DCD8D600DDD8D400AAA5
          A200E0DBD800C0BBB800C9C4C200A29D9A00BCB7B400A39D9D00D2CFD000FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00F1F1F100F2F3F300D0D0CF00B0AEAD00C8C6
          C500C0BEBD00D0CECD00C4C2C100BAB8B700C3C1C000C5C3C000CCCAC500B0AD
          A900CECCC900D3D1CF00CFCDCB00ABAAA900CBCCCB00C7C5C400BEBBB900BDBA
          B800BDBBB800BCB9B500CECAC600B4B0AB00B8B4AF00B5B0AD00A6A2A000FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00F7F7F700F4F4F400EDECEC00D7D5D300CECC
          CB00C9C7C600CCCAC900D4D2D100C2C0BF00B1AFAE00DEDBD600CBC7BD00C9C5
          BD00C0BDB700C6C4C000B7B8B500B8B9B800B9BCBC00C2C6C500C8C9C700BCBD
          BB00BAB9B500BFBEB900BAB7B000AEABA300B3B1A800C4BFB8008A847E00F4F3
          F200FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FCFCFC00F2F3F300F7F6F600DEDCDB00D4D2
          D100BAB8B700C8C6C500DCDAD900B7B5B400BAB9B800CBC8C300BDB9B000CDCB
          C200BEBBB500CDCAC600CACAC700BDBFBE00BBBDBC00C2C1BD00BBB8B200C2BE
          BA00ACA9A300D0CBC700E0DCD700CBC8C200A9A59F00C9C6C1008F8B8700CECD
          CB00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F0F0F000F9F9F800E2E0DF00ABA9
          A800BBB9B800C9C7C600D5D3D200CECCCB00D0CECD00E0DDD900DBD8D000C9C6
          C000C2BFBA00CECCC700C0C0BC00C7C7C400ACADAA00D4D1CC00E7E1DC00E3DF
          DC00B1AEAA00CFCBCA00CDCAC900D2D0CF00A9A7A800C2C0BD00A8A7A400A1A0
          9F00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F8F8F800F6F6F500EEEDEC00D6D4
          D300D4D2D100D8D6D500E8E7E600CDCBCA00D0CECD00CBC8C600DBD8D400D0CD
          C800C2C0BB00C7C7C200D7D6D200BEBCB800BBB9B500C7C6C400D4D3D200C9C8
          C900C2C2C400B5B6BA00C1C3C900C4C6CD00B9BBC200C1BFBC00C4C2BB007B79
          7700F6F5F500FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FDFDFD00F2F2F200FCFAF900DCDA
          D900D6D4D300CDCBCA00C8C6C500DFDDDC00D1CFCE00B7B5B400D6D6D200B4B4
          B000BDBCB800BEBDB800D1CFCA00BFBCB600C9C6C000C5C4C300C5C7C900C5C7
          C900BFC1C400BDC1C300BCC0C400B8BABF00BBBEC100C8C4B800D8CFC0008B83
          7C00DCDAD900FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F4F4F400FAF9F800ECEB
          EA00C1BFBE00BAB8B700C0BEBD00D8D6D500B8B5B400C6C5C400CBCDCC00C5C7
          C500D0D1CD00D1CFCA00D8D5D000D2D0C800D1CEC500D4D2CB00D0CFC800CDCA
          C300CDCAC000CEC9BE00D2CCBF00D6CEBF00D5CCBC00DBCDB200E2D2B600B5A6
          9600B2A9A600FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F9F9F900F7F5F400F7F5
          F300C3C1BF00C0BEBD00D2D0CF00E0DEDD00D6D3D300DCDDDD00E5E8EA00E2E4
          E500DDDFDD00E7E6E400E2E0DC00E3E2DB00DFDDD600DDD7C700E0D7BF00E0D6
          BC00DFD0B200E1CEAC00E0CBA300DDC49900DDC39600D4BB8E00CDB58F00E0CA
          B2008E7E7500FAF9F900FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FEFEFE00F5F2F500FCFA
          FD00EEECEE00EFEDEF00F7F5F400FBFBF800FBF8F300FBFAF200F5F4E800F5F0
          E300F2EBDB00EEE4D200EBE0CB00E0D3BC00DED0B800DDCAA800D6BE9400D3B7
          8B00D0B18200CAAA7900CBA77200C8A36C00BD996100BC9A6400B2936600E0C4
          A700A38C8100E7E2E100FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00F6F6F700F8F8
          F900FFFFFE00FFFDF400FEF9EA00FBF4DC00F8ECCF00F4E3BD00EFD8A900EACE
          9E00E4C69200DEBD8500DBB77A00D7B07100D1AA6700D1A66700D0A16500CA9A
          6000C7976200C3946300B98F62008F6B3F0088683E009B794600B6926400D3B2
          9400CCAFA100CCBEBD00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FDFEF400F2F0
          D400E1D8B200E0D1A000DCC89000D5BC8000D0B77800D0B37800CFAA7500CCA7
          7000CAA56B00C8A16500C9A06100C99F5E00C89D5B00C99B5B00C99A5C00C897
          5D00C7956000C3946300B88F610094714500B8986E009F7B4800BB956500BE9B
          7B00E1C4B400B3A09C00FDFDFD00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFC00F6EF
          D300CDB58500CFAA6C00D9AC6800D9AC6900D1A86900CEA56A00CDA26A00CCA1
          6800CB9F6800CA9E6700CAA06900C89F6600C59C6500C89D6400CC9F6500CD9F
          6800CB9D6900C99C6D00C69E7200A2805500AD8F6500A8835200AE895A00BD9B
          7A00DEC2B000BFABA400F0EDED00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFF6
          F300E5C2A900D5A27500DDA27000DBA17500D39F7C00D2A17A00D4A67100D5A7
          7200D5A97600D3A97700D1AA7B00D2A97D00D1AA7F00D2AC7C00D5B07C00D8B0
          7F00D6AD7F00D1AA7F00CFAB8300B69871009277510097784B00BA996F00C3A5
          8600CDB5A200CEBDB500DBD6D500FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFD
          FE00F0D8C600D4AC8A00DCAD8500DCAF8C00D8B09700DBB49600E2B98D00E1BA
          8F00E0BA9300DEBB9500DCBB9800DABA9A00D9BA9C00D8BC9800D9BF9700DDBF
          9A00DCBD9900D7B99700D2B69600CEB79800C8B49700CBB39400CFB79800D1BB
          A400D6C3B400D7CAC200C5C0BD00FBFAFA00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FE00F6F1D500DFCFA700E1CCA100E2CEA500DFCEA900E2CEAE00E8CCB100E8CC
          B100E4CBB200E4CCB400E2CCB600DFCAB700DDC9B700DDCDB800DED0B800E0D0
          B800E0CEB800E1CEBB00E0D0BC00DCD0BD00D9D0BD00E0D4C200E3D6C600E1D5
          C700DDD3C800E0D8D200B7B2AE00EBE9E900FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FBF9EC00E2DCC700DED6BE00E2DABD00E4DDBD00E3DAC300E2D6CB00E7D9
          CF00EADDD400E8DCD300E5D9CF00E4D9D000E4D9D000DFD9CF00DAD9CE00DBD6
          CB00DDD4CC00DED5CC00DAD6CD00D8D7CF00D6D7D000D5D4D100D4D1CF00D4D1
          CE00D5D4CE00DCDAD500C5C1BD00E1DFDD00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FEFBFD00E8E3E700E5E0DF00EAE4E000EFE7DE00F0EBE300EFECEA00F1ED
          EC00F4EFEE00F3EFED00EFEBE700EAE6E300E6E1DE00DBDCDA00D3D7D500D2D3
          D100D4D3D000D3D1CF00D1D1D000D0D3D300CFD5D400CCD1D400D1D4D800D7DA
          DB00D7D9D800DCDCD900ECEBE900FDFDFD00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FEFEFE00F3F2F300F5F4F500FAFAFA00FAF8F800FAF9F900FAFAFA00F8F8
          F800F7F6F600F4F3F300F0EFEF00EAE9E900E5E4E400DDDDDD00D8D8D800D8D8
          D800DADADB00DCDCDC00E2E2E300E7E8E800ECECED00EEEEEF00F3F4F400F5F6
          F600FAFAFA00FCFDFD00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00F7F7F700F3F3F300FBFBFB00F9F9F900F8F8F800F7F7F700F5F5
          F500F5F5F500F1F1F100EEEEEE00EFEFEF00EEEFEF00F1F1F100F4F4F400F7F7
          F700F7F7F700FBFBFB00FEFEFE00FEFEFE00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FDFDFD00F5F5F500F3F3F300F1F1F100F7F7F700FAFAFA00FBFB
          FB00FAFAFA00FAFAFA00FCFCFC00FEFEFE00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FDFDFD00FBFBFB00FCFCFC00FDFDFD00FEFEFE00FFFF
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
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00}
        GlyphMask.Data = {00000000}
        ParentFont = False
        OnClick = RecalculeDocumentClick
      end
      object LMONTANTRET: TFlashingLabel
        Left = 222
        Top = 99
        Width = 25
        Height = 13
        Caption = 'XXX'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object GP_MODEREGLE: THValComboBox
        Left = 287
        Top = 45
        Width = 87
        Height = 19
        Style = csSimple
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ItemHeight = 13
        ParentFont = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TTMODEREGLE'
      end
      object PTotaux1: THPanel
        Left = 574
        Top = 2
        Width = 163
        Height = 85
        BorderStyle = bsSingle
        Caption = ' '
        FullRepaint = False
        TabOrder = 5
        OnResize = PTotaux1Resize
        BackGroundEffect = bdFond
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        ColorNb = 2
        TextEffect = tenone
        object HGP_TOTALTTCDEV: THLabel
          Left = 2
          Top = 24
          Width = 30
          Height = 13
          Caption = 'T.T.C.'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object HGP_ACOMPTEDEV: THLabel
          Left = 2
          Top = 44
          Width = 72
          Height = 13
          Caption = 'Acompte/R'#232'gl.'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LGP_ACOMPTEDEV: THLabel
          Left = 128
          Top = 44
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object LGP_TOTALTTCDEV: THLabel
          Left = 128
          Top = 24
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALTAXESDEV: THLabel
          Left = 2
          Top = 4
          Width = 29
          Height = 13
          Caption = 'Taxes'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LTOTALTAXESDEV: THLabel
          Left = 128
          Top = 4
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object LNETAPAYERDEV: THLabel
          Left = 128
          Top = 64
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_NETAPAYERDEV: THLabel
          Left = 2
          Top = 64
          Width = 55
          Height = 13
          Caption = 'Net '#224' payer'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LHTTCDEV: THLabel
          Left = 128
          Top = 84
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
          Visible = False
        end
        object LCAUTION: TFlashingLabel
          Left = 2
          Top = 84
          Width = 97
          Height = 13
          Alignment = taRightJustify
          Caption = 'Caution bancaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
          Visible = False
        end
        object HGP_TOTALTIMBRES: THLabel
          Left = 2
          Top = 104
          Width = 82
          Height = 13
          Caption = 'Droits de Timbres'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LGP_TOTALTIMBRES: THLabel
          Left = 128
          Top = 104
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALRETENUTTC: THLabel
          Left = 2
          Top = 124
          Width = 68
          Height = 13
          Caption = 'Retenues Div.'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
          Visible = False
        end
        object LTOTALRETENUTTC: THLabel
          Left = 128
          Top = 124
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
      end
      object PTotaux: THPanel
        Left = 418
        Top = 2
        Width = 154
        Height = 85
        BorderStyle = bsSingle
        Caption = ' '
        FullRepaint = False
        TabOrder = 4
        OnResize = PTotauxResize
        BackGroundEffect = bdFond
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        ColorNb = 2
        TextEffect = tenone
        object HGP_TOTALHTDEV: THLabel
          Left = 2
          Top = 4
          Width = 21
          Height = 13
          Caption = 'H.T.'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LGP_TOTALHTDEV: THLabel
          Left = 119
          Top = 4
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object LGP_TOTALREMISEDEV: THLabel
          Left = 119
          Top = 64
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALREMISEDEV: THLabel
          Left = 2
          Top = 64
          Width = 35
          Height = 13
          Caption = 'Remise'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALPORTSDEV: THLabel
          Left = 2
          Top = 24
          Width = 71
          Height = 13
          Caption = 'dont ports/frais'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LTOTALPORTSDEV: THLabel
          Left = 119
          Top = 24
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALESCDEV: THLabel
          Left = 2
          Top = 44
          Width = 47
          Height = 13
          Caption = 'Escompte'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
        end
        object LGP_TOTALESCDEV: THLabel
          Left = 119
          Top = 44
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALRGDEV: THLabel
          Left = 2
          Top = 84
          Width = 22
          Height = 13
          Caption = 'R.G.'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
          Visible = False
        end
        object LTOTALRGDEV: THLabel
          Left = 119
          Top = 84
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object HGP_TOTALRETENUHT: THLabel
          Left = 2
          Top = 104
          Width = 68
          Height = 13
          Caption = 'Retenues Div.'
          Color = clInfoBk
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          Transparent = True
          Visible = False
        end
        object LTOTALRETENUHT: THLabel
          Left = 119
          Top = 104
          Width = 26
          Height = 13
          Alignment = taRightJustify
          Caption = '0,00'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
      end
      object GP_ESCOMPTE: THNumEdit
        Left = 287
        Top = 3
        Width = 87
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 0
        OnExit = GP_ESCOMPTEExit
        Decimals = 2
        Digits = 15
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object GP_REMISEPIED: THNumEdit
        Left = 287
        Top = 24
        Width = 87
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        OnExit = GP_REMISEPIEDExit
        Decimals = 2
        Digits = 15
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object INFOSLIGNE: THGrid
        Left = 2
        Top = 2
        Width = 216
        Height = 82
        ColCount = 2
        Ctl3D = True
        DefaultColWidth = 40
        DefaultRowHeight = 16
        FixedCols = 0
        RowCount = 2
        FixedRows = 0
        Options = [goRowSelect]
        ParentCtl3D = False
        ScrollBars = ssVertical
        TabOrder = 6
        SortedCol = -1
        Titres.Strings = (
          '')
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
      end
      object GP_TOTALHTDEV: THNumEdit
        Tag = 1
        Left = 42
        Top = 29
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object GP_TOTALREMISEDEV: THNumEdit
        Tag = 1
        Left = 75
        Top = 29
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object HTOTALTAXESDEV: THNumEdit
        Tag = 1
        Left = 108
        Top = 29
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object GP_TOTALESCDEV: THNumEdit
        Tag = 1
        Left = 140
        Top = 29
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object GP_TOTALTTCDEV: THNumEdit
        Tag = 1
        Left = 173
        Top = 29
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object GP_ACOMPTEDEV: THNumEdit
        Tag = 1
        Left = 10
        Top = 29
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TOTALPORTSDEV: THNumEdit
        Tag = 1
        Left = 10
        Top = 45
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 12
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NETAPAYERDEV: THNumEdit
        Tag = 1
        Left = 42
        Top = 49
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TOTALRGDEV: THNumEdit
        Tag = 1
        Left = 78
        Top = 49
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 14
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object HTTCDEV: THNumEdit
        Tag = 1
        Left = 114
        Top = 49
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 15
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TTYPERG: TEdit
        Left = 148
        Top = 48
        Width = 25
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 16
        Visible = False
        OnChange = AfficheZonePied
      end
      object TCaution: TEdit
        Left = 172
        Top = 48
        Width = 25
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 17
        Visible = False
        OnChange = AfficheZonePied
      end
      object GP_TOTALTIMBRES: THNumEdit
        Tag = 1
        Left = 11
        Top = 64
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 18
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TOTALRETENUHT: THNumEdit
        Tag = 1
        Left = 49
        Top = 66
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 19
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TOTALRETENUTTC: THNumEdit
        Tag = 1
        Left = 89
        Top = 66
        Width = 36
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 20
        Visible = False
        OnChange = AfficheZonePied
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
    object PButtons: TPanel
      Left = 0
      Top = 86
      Width = 995
      Height = 36
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 0
      object DockBottom: TDock97
        Left = 0
        Top = 5
        Width = 995
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
          DesignSize = (
            906
            27)
          object BMenuZoom: TToolbarButton97
            Tag = -100
            Left = 28
            Top = 0
            Width = 40
            Height = 27
            Hint = 'Menu zoom'
            Caption = 'Zoom'
            DisplayMode = dmGlyphOnly
            DropdownArrow = True
            DropdownMenu = POPZ
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnMouseEnter = BMenuZoomMouseEnter
            GlobalIndexImage = 'Z0061_S16G1'
            IsControl = True
          end
          object BEche: TToolbarButton97
            Tag = 1
            Left = 108
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Ech'#233'ances'
            Caption = 'Ech'#233'ances'
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            OnClick = BEcheClick
            GlobalIndexImage = 'Z0041_S16G2'
          end
          object BInfos: TToolbarButton97
            Tag = 1
            Left = 68
            Top = 0
            Width = 40
            Height = 27
            Hint = 'Actions compl'#233'mentaires'
            Caption = 'Actions'
            DisplayMode = dmGlyphOnly
            DropdownArrow = True
            DropdownMenu = POPY
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            NumGlyphs = 2
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BInfosClick
            GlobalIndexImage = 'Z0105_S16G2'
            IsControl = True
          end
          object BChercher: TToolbarButton97
            Tag = 1
            Left = 340
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Rechercher dans la pi'#232'ce'
            Caption = 'Rechercher'
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BChercherClick
            GlobalIndexImage = 'Z0077_S16G1'
            IsControl = True
          end
          object BSousTotal: TToolbarButton97
            Tag = 1
            Left = 228
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Ins'#233'rer un sous-total'
            Caption = 'Sous-total'
            AllowAllUp = True
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BSousTotalClick
            GlobalIndexImage = 'Z0797_S16G1'
            IsControl = True
          end
          object BAcompte: TToolbarButton97
            Tag = 1
            Left = 256
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Acomptes/R'#232'glements'
            Caption = 'Acomptes/R'#232'glements'
            AllowAllUp = True
            DisplayMode = dmGlyphOnly
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BAcompteClick
            GlobalIndexImage = 'Z0040_S16G1'
            IsControl = True
          end
          object BActionsLignes: TToolbarButton97
            Tag = 1
            Left = 182
            Top = 0
            Width = 40
            Height = 27
            Hint = 'Actions lignes'
            Caption = 'Actions'
            DisplayMode = dmGlyphOnly
            DropdownArrow = True
            DropdownMenu = PopL
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            Layout = blGlyphTop
            Margin = 2
            NumGlyphs = 2
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            GlobalIndexImage = 'Z0543_S16G2'
            IsControl = True
          end
          object Sep1: TToolbarSep97
            Left = 222
            Top = 0
          end
          object Sep2: TToolbarSep97
            Left = 176
            Top = 0
          end
          object BOffice: TToolbarButton97
            Tag = 1
            Left = 508
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Edition bureautique'
            Caption = 'Bureautique'
            AllowAllUp = True
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            Visible = False
            OnClick = BOfficeClick
            GlobalIndexImage = 'Z0252_S16G1'
            IsControl = True
          end
          object BVentil: TToolbarButton97
            Tag = 1
            Left = 136
            Top = 0
            Width = 40
            Height = 27
            Hint = 'Ventilations analytiques'
            Caption = 'Analytiques'
            DisplayMode = dmGlyphOnly
            DropdownArrow = True
            DropdownMenu = PopV
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            Layout = blGlyphTop
            Margin = 2
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            GlobalIndexImage = 'Z0133_S16G1'
            IsControl = True
          end
          object BDescriptif: TToolbarButton97
            Tag = 1
            Left = 424
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Descriptif d'#233'taill'#233' de l'#39'article'
            Caption = 'Descriptif'
            AllowAllUp = True
            GroupIndex = 1
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BDescriptifClick
            GlobalIndexImage = 'Z0029_S16G1'
            IsControl = True
          end
          object BDelete: TToolbarButton97
            Left = 368
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Supprimer la pi'#232'ce'
            Caption = 'Supprimer'
            DisplayMode = dmGlyphOnly
            Layout = blGlyphTop
            Opaque = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            OnClick = BDeleteClick
            GlobalIndexImage = 'Z0005_S16G1'
          end
          object BImprimer: TToolbarButton97
            Left = 536
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Imprimer'
            Caption = 'Imprimer'
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            Layout = blGlyphTop
            Margin = 2
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            OnClick = BImprimerClick
            GlobalIndexImage = 'Z0369_S16G1'
          end
          object BPorcs: TToolbarButton97
            Tag = 1
            Left = 284
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Ports et frais'
            Caption = 'Acompte'
            AllowAllUp = True
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BPorcsClick
            GlobalIndexImage = 'Z0661_S16G1'
            IsControl = True
          end
          object BSaisieAveugle: TToolbarButton97
            Left = 605
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Saisie par codes barres'
            Caption = 'Saisie code barre'
            DisplayMode = dmGlyphOnly
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            Layout = blGlyphTop
            Margin = 4
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            OnClick = BSaisieAveugleClick
            GlobalIndexImage = 'Z1334_S16G1'
          end
          object BNouvelArticle: TToolbarButton97
            Left = 672
            Top = 0
            Width = 29
            Height = 27
            Hint = 'Nouvel article'
            Caption = 'Nouvel article'
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            Layout = blGlyphTop
            Margin = 2
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            OnClick = BNouvelArticleClick
            GlobalIndexImage = 'Z0053_S16G1'
          end
          object BRetenuGar: TToolbarButton97
            Tag = 1
            Left = 312
            Top = 0
            Width = 28
            Height = 27
            Hint = 'retenue de garantie'
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            Visible = False
            OnClick = BRetenuGarClick
            GlobalIndexImage = 'Z0109_S16G1'
          end
          object BPrixMarche: TToolbarButton97
            Left = 701
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Application d'#39'un prix marche'
            Caption = 'Application d'#39'un prix marche'
            Anchors = [akRight, akBottom]
            DisplayMode = dmGlyphOnly
            Opaque = False
            ParentShowHint = False
            ShowHint = True
            OnClick = BPrixMarcheClick
            GlobalIndexImage = 'Z0143_S16G1'
          end
          object BArborescence: TToolbarButton97
            Tag = 1
            Left = 396
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Liste des paragraphes du document'
            Caption = 'Arborescence Paragraphes'
            AllowAllUp = True
            GroupIndex = 1
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BArborescenceClick
            GlobalIndexImage = 'Z0189_S16G0'
            IsControl = True
          end
          object BtQteAuto: TToolbarButton97
            Left = 633
            Top = 0
            Width = 39
            Height = 27
            Hint = 'Affecte les quantit'#233's automatiquement'
            Caption = 'Affecte les quantit'#233's automatiquement'
            DisplayMode = dmGlyphOnly
            DropdownArrow = True
            DropdownMenu = PopA
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'MS Sans Serif'
            Font.Style = [fsBold]
            Layout = blGlyphTop
            Margin = 4
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            GlobalIndexImage = 'Z0213_S16G1'
          end
          object Bminute: TToolbarButton97
            Left = 729
            Top = 0
            Width = 27
            Height = 27
            Hint = 'Impression de la minute'
            DisplayMode = dmGlyphOnly
            Opaque = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            OnClick = BminuteClick
            GlobalIndexImage = 'M0015_S16G1'
          end
          object RecalculeDocument: TToolbarButton97
            Tag = 1
            Left = 480
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Recalcul du document'
            Caption = 'Recalcul du document'
            AllowAllUp = True
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = RecalculeDocumentClick
            GlobalIndexImage = 'Z0956_S16G1'
            IsControl = True
          end
          object BCalculDocAuto: TToolbarButton97
            Tag = 1
            Left = 452
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Calcul Automatique'
            Caption = 'Calcul Automatique'
            AllowAllUp = True
            GroupIndex = 2
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BCalculDocAutoClick
            GlobalIndexImage = 'M0066_S16G1'
            IsControl = True
          end
          object BintegreExcel: TToolbarButton97
            Left = 564
            Top = 0
            Width = 41
            Height = 27
            Hint = 'R'#233'cup'#233'ration de lignes via Excel'
            DropdownAlways = True
            DropdownArrow = True
            DropdownCombo = True
            DropdownMenu = PopIntExcel
            Glyph.Data = {
              36040000424D3604000000000000360000002800000010000000100000000100
              2000000000000004000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000082
              19FF008219FF0000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000008219FF0FF2
              98FF0FF298FF008219FF00000000000000000000000000000000000000000000
              00007F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF008219FF0FF2
              98FF0FF298FF008219FF00000000000000000000000000000000000000007F7F
              7FFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFF008219FF0FF2
              98FF0FF298FF008219FF7F7F7FFF000000000000000000000000000000007F7F
              7FFFDADADAFFDADADAFFDADADAFF008219FF008219FF008219FF008219FF0FF2
              98FF0FF298FF008219FF008219FF008219FF008219FF00000000000000007F7F
              7FFFEEEEEEFFEEEEEEFF008219FF0FF298FF0FF298FF0FF298FF0FF298FF0FF2
              98FF0FF298FF0FF298FF0FF298FF0FF298FF0FF298FF008219FF000000000000
              00007F7F7FFF7F7F7FFF008219FFABF3CFFFABF3CFFFABF3CFFFABF3CFFF0FF2
              98FF0FF298FFABF3CFFFABF3CFFFABF3CFFFABF3CFFF008219FF000000007F7F
              7FFFDADADAFFDADADAFFDADADAFF008219FF008219FF008219FF008219FF0FF2
              98FF0FF298FF008219FF008219FF008219FF008219FF00000000000000007F7F
              7FFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFF008219FF0FF2
              98FF0FF298FF008219FF7F7F7FFF000000000000000000000000000000007F7F
              7FFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFF008219FF0FF2
              98FF0FF298FF008219FF7F7F7FFF000000000000000000000000000000000000
              00007F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF008219FFABF3
              CFFFABF3CFFF008219FF00000000000000000000000000000000000000007F7F
              7FFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFF0082
              19FF008219FFDADADAFF7F7F7FFF000000000000000000000000000000007F7F
              7FFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADADAFFDADA
              DAFFDADADAFFDADADAFF7F7F7FFF000000000000000000000000000000007F7F
              7FFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFFEEEEEEFFEEEE
              EEFFEEEEEEFFEEEEEEFF7F7F7FFF000000000000000000000000000000000000
              00007F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F7FFF7F7F
              7FFF7F7F7FFF7F7F7FFF00000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000}
            GlyphMask.Data = {00000000}
            ParentShowHint = False
            ShowHint = True
            Visible = False
          end
          object BMinute2: TToolbarButton97
            Left = 756
            Top = 0
            Width = 27
            Height = 27
            Hint = 'Tableau minute'
            DisplayMode = dmGlyphOnly
            Opaque = False
            ParentShowHint = False
            ShowHint = True
            Visible = False
            OnClick = Bminute2Click
            GlobalIndexImage = 'Z0445_S16G1'
          end
          object BDEMPRIX: TToolbarButton97
            Left = 783
            Top = 0
            Width = 39
            Height = 27
            Hint = 'Gestion des demandes de prix'
            DropdownArrow = True
            DropdownCombo = True
            DropdownMenu = POPDEMPRIX
            ParentShowHint = False
            ShowHint = True
            GlobalIndexImage = 'Z2213_S16G1'
          end
          object BRENUM: TToolbarButton97
            Left = 865
            Top = 0
            Width = 41
            Height = 27
            Hint = 'Num'#233'rotation des lignes'
            DropdownArrow = True
            DropdownMenu = PopNumauto
            ParentShowHint = False
            ShowHint = True
            GlobalIndexImage = 'Z0122_S24G1'
          end
          object TBVOIRTOT: TToolbarButton97
            Left = 0
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Voir les totalisations documents'
            Glyph.Data = {00000000}
            GlyphMask.Data = {00000000}
            ImageIndex = 0
            Images = HimgTOT
            Opaque = False
            ParentShowHint = False
            ShowHint = True
            OnClick = TBVOIRTOTClick
          end
          object BGED: TToolbarButton97
            Left = 822
            Top = 0
            Width = 43
            Height = 27
            Hint = 'Num'#233'rotation des lignes'
            DropdownAlways = True
            DropdownArrow = True
            DropdownCombo = True
            DropdownMenu = POPGEDBSV
            Glyph.Data = {
              F6060000424DF606000000000000360000002800000018000000180000000100
              180000000000C0060000120B0000120B00000000000000000000FFFFFFFFFEFE
              6B92DB6E94DCCAD9EFF9FBFBFFFFFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFEFEFF2002EBD003AD10040C93A67D1A1B9E5DCE7F3FAFBFDFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFDBDDE2001AA50140D90038CE0026CA0021C800
              30C80A41C37D9BD9CCDBF0F4F6FBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC8CBCE0E49B00037D8033E
              CD0442CE0236CB001DC20013BF000EBD0021BA023ABA3D67C697ABDBE3EBF3FF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFB7BBC1
              4066B40030D3033ECD0035CC0C43CDBACEEDC3D4F02252C6002CBB0021BE0009
              B0001AB00028B02049B67592CDBFCCE8EDF0F8FFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFF9BA2B16681B90031D1033ECD0021C582A5E5FEFEFC5783D6C1D5EE
              6F8FD3073DBD638CD51D4CBF000FA70016AC000CA400159E002AA23A5CB4A4B8
              DBEEF2F7FFFFFFFFFFFFFFFFFF79899B8E9CC50035D0013CCD001FC4AECCEDAA
              ADE80001B5B4D3E8BBC5EA3266C8FFFFFECADBEAD5DDED2D61BB023FAE0021A3
              000A98001EA0001C9900309AE7EAF3FFFFFFFCFDFD46597DBBC4D30043D00038
              CE0024C7749DE3BEC7EC002BBFA4B6E687A2DF4472CDD4DBEC0019A4315BC431
              5AB8DAE5EECEDBEA89A4D400179600279E002398BFCEE4FFFFFFF4F5F7364877
              E4E4DF144FD00035CD0038CD033ACADFEEF44467D00000B01B4FC52251C4F7FA
              FAAABDE45073C90013A4DAE2F0577CC0BDC7E4B8C4DD00139100239998AAD3FF
              FFFFE1E4EB334B7AFFFEEA2B5DD4002FCC0340CE0027C72C56CDD3DDF1F1F6F8
              C6D2EC0029B4D2DBEF4A69C7597BCD001DA5D7E5F15B6DC2000085D6DFED6180
              C100068C768EC6FFFFFFD4D8E42A4576FFFFF2456ED40029CC043DCD0344CE00
              24C6143DC76E9ADE1943C40012B0C7D6E8B7C6E72957C00010A1C8D4EC7D8ECE
              0000858083C4CFE0EC0004893257ABFFFFFFB0BECD3B5282FFFFF96B85D40021
              C8043ECD033ECD023ECC0133C8001FBE0130C1002FBF0B43B88CAADDE4E8F366
              8FCA93AAD9D1D9ED001A9AA3B7DADBE3EF000F8E042F95F8FAFD95A3BD52678F
              FFFFFB9EADDE000FC3003CCE033FCE023BCB0138C8013AC6033DC30334C10029
              BB000CAB0012A8033BB31D50B38DA4D4C7D5E6FFFFFE375BAF001A95002C96E5
              EBF47B89AD647596FFFFF7DADFE91553CA002AC5001CC5002BC90031C80039C6
              023CC20331BF0231BC0132B80229B3002BB10121A80008990A34A70C30A2001C
              960035A1002393D4DAEB566E977F8BA8FBF7F0FBFAF4FFFBF0BEC8DF668BD12B
              60CE0034C30021BE023BC30332C00233BB0335B7032BB3012AB1012DAE0231AB
              002AA60028A20233A100319E002398ACBBD91C3776BDC8D1FFFFFBF5F1EBF5F2
              ECFFFDF3FFFFF6FFFDEFE5E6E893A7D60022BB0027C00233BC0334B7032BB302
              2CB10230AE0230AA0230A70131A60232A100319E001F96909ECC4764972D4A8C
              8392B6BBC1CCEEE6E0FFFFEEFDFBEEF5F1ECFFFFF0FFFFF53664C9001EB80010
              B0001DB20021B10029B1012FAE0330AA0230A70131A60232A100319E001F946B
              80C0FCFCFD8693B5324E8F062C7F183F894F6EA496A3C2D3D6D6FAF5E4FFFFF0
              FFFFF2D3D6E27F9BCE3B5FC3073FB60014A70009A0001AA3002DA80031A60233
              A100319E0021964263AFFFFFFFFFFFFFFFFFFFFCFDFDD3DBE693A4C94E67AA0F
              32912D4FA25571ADA5B0C4FEF7E7FFFFF4FFFFF4F4F6E9B7C1DD7E96CF3261B8
              00279F000F98001093001F97002097163C9DFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFF6F7FAC3CADF7C93C44158A5284A9C3E5DA47485B5D1CED1FF
              FCE9FFFFFAFFFFF5D9DDDF9AA9D1607DBF244FAA002193042993FFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F9FBCCD3
              E5415CA700128D19449B5470A8A3AAC4EEE4DDFFFFEFFFFDEBFCF9EECDCEDBEE
              F0F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFDEE2E9000887000C8D0003880008890029944061A5E8E8
              E7F3F1EEFFFFFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE1E8EEA9B5D3677DB916429C
              000D8A081D8FFDFEFDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFF4F6F8AEBACF858DBBFFFFFFFFFFFFFFFFFFFFFFFF}
            GlyphMask.Data = {00000000}
            ParentShowHint = False
            ShowHint = True
          end
        end
        object Valide97: TToolbar97
          Left = 915
          Top = 0
          Caption = 'Validation'
          CloseButton = False
          DefaultDock = DockBottom
          DockPos = 915
          TabOrder = 1
          DesignSize = (
            84
            27)
          object BValider: TToolbarButton97
            Left = 0
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Enregistrer la pi'#232'ce'
            Caption = 'Enregistrer'
            Anchors = [akTop, akRight]
            DisplayMode = dmGlyphOnly
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
            Caption = 'Annuler'
            Anchors = [akTop, akRight]
            Cancel = True
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ModalResult = 2
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BAbandonClick
            GlobalIndexImage = 'Z0021_S16G1'
            IsControl = True
          end
          object BAide: TToolbarButton97
            Tag = 1
            Left = 56
            Top = 0
            Width = 28
            Height = 27
            Hint = 'Aide'
            HelpContext = 119000017
            Caption = 'Aide'
            Anchors = [akTop, akRight]
            DisplayMode = dmGlyphOnly
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            Layout = blGlyphTop
            Opaque = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            Spacing = -1
            OnClick = BAideClick
            GlobalIndexImage = 'Z1117_S16G1'
            IsControl = True
          end
        end
      end
    end
  end
  object TTVParag: TToolWindow97
    Left = 693
    Top = 86
    ClientHeight = 385
    ClientWidth = 217
    Caption = 'Liste des Paragraphes'
    ClientAreaHeight = 385
    ClientAreaWidth = 217
    TabOrder = 17
    Visible = False
    OnClose = TTVParagClose
    object TVParag: TTreeView
      Left = 0
      Top = 0
      Width = 217
      Height = 385
      Align = alClient
      Color = clInfoBk
      HideSelection = False
      HotTrack = True
      Indent = 19
      ReadOnly = True
      TabOrder = 0
      ToolTips = False
      OnClick = TVParagClick
    end
  end
  object Descriptif1: THRichEditOLE
    Left = 144
    Top = 179
    Width = 229
    Height = 94
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    ScrollBars = ssVertical
    TabOrder = 22
    Visible = False
    WantTabs = True
    OnKeyDown = Descriptif1KeyDown
    Margins.Top = 0
    Margins.Bottom = 0
    Margins.Left = 0
    Margins.Right = 0
    ContainerName = 'Document'
    ObjectMenuPrefix = '&Object'
    LinesRTF.Strings = (
      
        '{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1036{\fonttbl{\f' +
        '0\fnil Times New Roman;}}'
      '{\*\generator Riched20 6.3.9600}\viewkind4\uc1 '
      '\pard\f0\fs20 '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par '
      '\par }')
  end
  object BZoomRessource: THBitBtn
    Tag = 100
    Left = 200
    Top = 152
    Width = 25
    Height = 29
    Hint = 'Voir commercial'
    Caption = 'Ressource'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 25
    Visible = False
    OnClick = BZoomCommercialClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object PGS: TPanel
    Left = 0
    Top = 125
    Width = 995
    Height = 388
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    TabStop = True
    object GS: THGrid
      Tag = 1
      Left = 0
      Top = 0
      Width = 995
      Height = 388
      Align = alClient
      BorderStyle = bsNone
      ColCount = 8
      Ctl3D = True
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 60
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goColSizing, goTabs]
      ParentCtl3D = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      OnExit = GSExit
      OnKeyDown = GSKeyDown
      OnMouseDown = GSMouseDown
      OnMouseMove = GSMouseMove
      OnSelectCell = GSSelectCell
      OnTopLeftChanged = GSTopLeftChanged
      SortedCol = -1
      Couleur = False
      MultiSelect = True
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GSRowEnter
      OnRowExit = GSRowExit
      OnCellEnter = GSCellEnter
      OnCellExit = GSCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
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
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
  end
  object PEnteteAffaire: THPanel
    Left = 93
    Top = 188
    Width = 728
    Height = 159
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Caption = ' '
    FullRepaint = False
    TabOrder = 1
    Visible = False
    OnExit = PEnteteExit
    BackGroundEffect = bdFond
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    ColorNb = 2
    TextEffect = tenone
    object HGP_NUMPIECE__: THLabel
      Left = 7
      Top = 28
      Width = 12
      Height = 13
      Caption = 'N'#176
      Transparent = True
    end
    object HGP_REFINTERNE__: THLabel
      Left = 250
      Top = 52
      Width = 53
      Height = 13
      Caption = 'Ref Interne'
      FocusControl = GP_REFINTERNE__
      Transparent = True
    end
    object HGP_TIERS__: THLabel
      Left = 250
      Top = 7
      Width = 26
      Height = 13
      Caption = 'Client'
      FocusControl = GP_TIERS__
      Transparent = True
    end
    object HGP_DATEPIECE__: THLabel
      Left = 120
      Top = 28
      Width = 12
      Height = 13
      Caption = 'du'
      FocusControl = GP_DATEPIECE__
      Transparent = True
    end
    object HGP_REPRESENTANT__: THLabel
      Left = 479
      Top = 52
      Width = 54
      Height = 13
      Caption = 'Commercial'
      FocusControl = GP_REPRESENTANT__
      Transparent = True
    end
    object HGP_ETABLISSEMENT__: THLabel
      Left = 479
      Top = 7
      Width = 65
      Height = 13
      Caption = 'Etablissement'
      FocusControl = GP_ETABLISSEMENT__
      Transparent = True
    end
    object FTitrePiece__: THLabel
      Left = 1
      Top = 2
      Width = 242
      Height = 20
      Alignment = taCenter
      AutoSize = False
      Caption = 'PIECE COMMERCIALE'
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
    object HGP_DEPOT__: THLabel
      Left = 479
      Top = 74
      Width = 29
      Height = 13
      Caption = 'D'#233'p'#244't'
      FocusControl = GP_DEPOT__
      Transparent = True
    end
    object BRechAffaire__: TToolbarButton97
      Left = 288
      Top = 26
      Width = 23
      Height = 22
      Hint = 'Rechercher client/affaire'
      Opaque = False
      OnClick = BRechAffaire__Click
      GlobalIndexImage = 'Z0002_S16G1'
    end
    object HGP_AFFAIRE__: THLabel
      Left = 250
      Top = 30
      Width = 30
      Height = 13
      Caption = 'Affaire'
      Transparent = True
    end
    object BDevise__: TToolbarButton97
      Tag = -100
      Left = 468
      Top = 26
      Width = 61
      Height = 21
      Caption = 'Devise'
      Alignment = taLeftJustify
      DisplayMode = dmTextOnly
      DropdownArrow = True
      DropdownMenu = PopD
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Layout = blGlyphTop
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      IsControl = True
    end
    object ISigneEuro__: TImage
      Left = 531
      Top = 29
      Width = 16
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        0000100000000100040000000000800000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
        FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
        4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
        FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
        4FF4}
      Stretch = True
      Transparent = True
      Visible = False
    end
    object BRazAffaire: TToolbarButton97
      Left = 288
      Top = 3
      Width = 23
      Height = 22
      Hint = 'Effacer s'#233'lection client/affaire'
      Opaque = False
      OnClick = BRazAffaireClick
      GlobalIndexImage = 'Z0080_S16G1'
    end
    object LibComplAffaire: THLabel
      Left = 5
      Top = 30
      Width = 207
      Height = 13
      AutoSize = False
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      Transparent = True
    end
    object HGP_DOMAINE__: THLabel
      Left = 479
      Top = 98
      Width = 71
      Height = 13
      Caption = 'Domaine activ.'
      FocusControl = GP_DOMAINE__
      Transparent = True
    end
    object HGP_DATELIVRAISON__: THLabel
      Left = 250
      Top = 98
      Width = 53
      Height = 13
      Caption = 'Livraison le'
      FocusControl = GP_DATELIVRAISON__
      Transparent = True
    end
    object HGP_REFEXTERNE__: THLabel
      Left = 250
      Top = 74
      Width = 56
      Height = 13
      Caption = 'Ref Externe'
      FocusControl = GP_REFEXTERNE__
      Transparent = True
    end
    object LblCommercial: TLabel
      Left = 622
      Top = 51
      Width = 68
      Height = 13
      Caption = 'LblCommercial'
      Visible = False
    end
    object LMATERIEL: TLabel
      Left = 16
      Top = 74
      Width = 37
      Height = 13
      Caption = 'Mat'#233'riel'
      Visible = False
    end
    object LBTETAT: TLabel
      Left = 16
      Top = 98
      Width = 64
      Height = 13
      Caption = 'Type d'#39'action'
      Visible = False
    end
    object LATTACHEMENT: TLabel
      Left = 243
      Top = 120
      Width = 60
      Height = 13
      Alignment = taRightJustify
      Caption = 'Attachement'
      Visible = False
    end
    object BATTACHEMENT: TToolbarButton97
      Left = 491
      Top = 114
      Width = 24
      Height = 24
      Visible = False
      GlobalIndexImage = 'O0009_S24G1'
    end
    object DELATTACHEMENT: TToolbarButton97
      Left = 518
      Top = 114
      Width = 24
      Height = 24
      Visible = False
      GlobalIndexImage = 'Z0080_S24G1'
    end
    object GP_REFEXTERNE__: TEdit
      Left = 314
      Top = 70
      Width = 137
      Height = 21
      MaxLength = 35
      TabOrder = 7
    end
    object GP_REFINTERNE__: TEdit
      Left = 314
      Top = 48
      Width = 137
      Height = 21
      MaxLength = 35
      TabOrder = 6
    end
    object GP_TIERS__: THCritMaskEdit
      Left = 314
      Top = 3
      Width = 136
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
      OnDblClick = GP_TIERSDblClick
      OnEnter = GP_TIERSEnter
      OnExit = GP_TIERSExit
      TagDispatch = 0
      DataType = 'GCTIERSSAISIE'
      ElipsisButton = True
      OnElipsisClick = GP_TIERSElipsisClick
    end
    object GP_DATEPIECE__: THCritMaskEdit
      Left = 137
      Top = 24
      Width = 78
      Height = 21
      AutoSelect = False
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 0
      Text = '  /  /    '
      OnExit = GP_DATEPIECEExit
      TagDispatch = 0
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object GP_DEVISE__: THValComboBox
      Left = 550
      Top = 26
      Width = 137
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 23
      OnChange = GP_DEVISEChange
      TagDispatch = 0
      DataType = 'TTDEVISE'
    end
    object GP_REPRESENTANT__: THCritMaskEdit
      Left = 550
      Top = 48
      Width = 71
      Height = 21
      TabOrder = 12
      OnExit = GP_REPRESENTANTExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GP_REPRESENTANTElipsisClick
    end
    object GP_NUMEROPIECE__: THPanel
      Left = 35
      Top = 24
      Width = 80
      Height = 21
      BevelOuter = bvLowered
      Caption = 'Non affect'#233
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 20
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object LIBELLETIERS__: THPanel
      Left = 7
      Top = 49
      Width = 218
      Height = 16
      Alignment = taLeftJustify
      BevelOuter = bvNone
      Ctl3D = True
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 21
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object GP_ETABLISSEMENT__: THValComboBox
      Left = 550
      Top = 3
      Width = 137
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 9
      OnChange = GP_ETABLISSEMENTChange
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
    end
    object GP_DEPOT__: THValComboBox
      Left = 550
      Top = 70
      Width = 137
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 13
      OnChange = GP_DEPOTChange
      OnExit = GP_DEPOT__Exit
      TagDispatch = 0
      DataType = 'GCDEPOT'
    end
    object GP_AFFAIRE1__: THCritMaskEdit
      Left = 314
      Top = 26
      Width = 42
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 2
      OnChange = GP_AFFAIRE1__Change
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AFFAIRE2__: THCritMaskEdit
      Left = 356
      Top = 26
      Width = 51
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 3
      OnChange = GP_AFFAIRE2__Change
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AFFAIRE3__: THCritMaskEdit
      Left = 408
      Top = 26
      Width = 43
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 4
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AFFAIRE: THCritMaskEdit
      Left = 416
      Top = 94
      Width = 25
      Height = 21
      TabStop = False
      Color = clYellow
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 19
      Visible = False
      TagDispatch = 0
    end
    object LDevise__: THPanel
      Left = 532
      Top = 28
      Width = 18
      Height = 16
      BevelOuter = bvNone
      Caption = '$'
      Ctl3D = True
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 24
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object GP_AVENANT__: THCritMaskEdit
      Left = 450
      Top = 26
      Width = 18
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 5
      OnExit = OnExitPartieAffaire
      TagDispatch = 0
    end
    object GP_AFFAIRE0__: THCritMaskEdit
      Left = 230
      Top = 26
      Width = 15
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      TabOrder = 18
      Visible = False
      TagDispatch = 0
    end
    object EuroPivot__: TEdit
      Left = 550
      Top = 26
      Width = 137
      Height = 21
      Enabled = False
      TabOrder = 10
      Visible = False
    end
    object GP_DOMAINE__: THValComboBox
      Left = 551
      Top = 94
      Width = 137
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 14
      OnExit = GP_DOMAINE__Exit
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'BTDOMAINEACT'
    end
    object GP_DATELIVRAISON__: THCritMaskEdit
      Left = 314
      Top = 94
      Width = 80
      Height = 21
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 8
      Text = '  /  /    '
      OnExit = GP_DATELIVRAISONExit
      TagDispatch = 0
      Operateur = Egal
      OpeType = otDate
      ElipsisButton = True
      ElipsisAutoHide = True
      ControlerDate = True
    end
    object LIBELLEAFFAIRE: THPanel
      Left = 7
      Top = 138
      Width = 682
      Height = 16
      Alignment = taLeftJustify
      BevelOuter = bvNone
      Ctl3D = True
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 22
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
    end
    object GP_RESSOURCE: THCritMaskEdit
      Left = 549
      Top = 47
      Width = 72
      Height = 21
      Hint = 'Ressource achat'
      Color = clWhite
      TabOrder = 11
      OnExit = GP_RESSOURCEExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GP_RESSOURCEElipsisClick
    end
    object GP_CODEMATERIEL: THCritMaskEdit
      Left = 86
      Top = 70
      Width = 147
      Height = 21
      TabOrder = 16
      Visible = False
      OnExit = GP_CODEMATERIELExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GP_CODEMATERIELElipsisClick
    end
    object GP_BTETAT: THCritMaskEdit
      Left = 86
      Top = 94
      Width = 147
      Height = 21
      TabOrder = 17
      Visible = False
      OnExit = GP_BTETATExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GP_BTETATElipsisClick
    end
    object ATTACHEMENT: THCritMaskEdit
      Left = 314
      Top = 116
      Width = 175
      Height = 21
      Enabled = False
      TabOrder = 15
      Visible = False
      OnExit = ATTACHEMENTExit
      TagDispatch = 0
    end
  end
  object TDescriptif: TToolWindow97
    Left = 384
    Top = 250
    ClientHeight = 259
    ClientWidth = 315
    Caption = 'Descriptif d'#233'taill'#233
    ClientAreaHeight = 259
    ClientAreaWidth = 315
    TabOrder = 14
    Visible = False
    OnClose = TDescriptifClose
    object Descriptif: THRichEditOLE
      Left = 0
      Top = 0
      Width = 315
      Height = 259
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnChange = DescriptifChange
      OnKeyDown = DescriptifKeyDown
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1036{\fonttbl{\f' +
          '0\fnil Times New Roman;}}'
        '{\*\generator Riched20 6.3.9600}\viewkind4\uc1 '
        '\pard\f0\fs20 '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par }')
    end
  end
  object BZoomAffaire: THBitBtn
    Tag = 100
    Left = 227
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Voir affaire'
    Caption = 'Affaire'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
    Visible = False
    OnClick = BZoomAffaireClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomPrecedente: THBitBtn
    Tag = 100
    Left = 162
    Top = 139
    Width = 28
    Height = 27
    Hint = 'Pi'#232'ce pr'#233'c'#233'dente'
    Caption = 'Pr'#233'c'#233'dente'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Visible = False
    OnClick = BZoomPrecedenteClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomCatalog: THBitBtn
    Tag = 100
    Left = 282
    Top = 131
    Width = 28
    Height = 27
    Hint = 'Catalogue fournisseur'
    Caption = 'Catalogue fournisseur'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 27
    Visible = False
    OnClick = BZoomCatalogClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 620
    Top = 176
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
      'ATTENTION : Pi'#232'ce non enregistr'#233'e !'
      
        'ATTENTION : Cette pi'#232'ce en cours de traitement par un autre util' +
        'isateur n'#39'a pas '#233't'#233' enregistr'#233'e !'
      'Liste des commerciaux'
      'Conditionnements'
      'Conditionnement en cours :'
      'Non affect'#233
      'Liste des commerciaux'
      'Liste des d'#233'p'#244'ts'
      
        'ATTENTION : Cette pi'#232'ce ne peut pas passer en comptabilit'#233' et n'#39 +
        'a pas '#233't'#233' enregistr'#233'e !'
      'Euro'
      'Autres'
      'Libell'#233's automatiques'
      '(r'#233'f'#233'rence de substitution possible :'
      '(remplacement par r'#233'f'#233'rence :'
      'ATTENTION : l'#39'impression ne s'#39'est pas correctement effectu'#233'e !'
      'ATTENTION : la suppression ne s'#39'est pas correctement effectu'#233'e !'
      'Cr'#233'dit accord'#233' : '
      'Encours actuel :'
      
        'ATTENTION : La pi'#232'ce pr'#233'sente un probl'#232'me de num'#233'rotation et n'#39'a' +
        ' pas '#233't'#233' enregistr'#233'e !'
      'Client'
      'Fournisseur'
      '26;'
      
        'ATTENTION. Le stock disponible est insuffisant pour certains art' +
        'icles.'
      'Changement de code TVA'
      ' ')
    Left = 489
    Top = 168
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 103
    Top = 150
  end
  object HPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Il n'#39'existe pas de tarif pour cet article correspond' +
        'ant '#224' la devise de la pi'#232'ce !;E;O;O;O;'
      
        '1;?caption?;Saisie impossible. Il n'#39'existe pas de tarif pour cet' +
        ' article correspondant '#224' la devise de la pi'#232'ce !;E;O;O;O;'
      
        '2;?caption?;Saisie impossible. Cet article ferm'#233' n'#39'est pas autor' +
        'is'#233' pour cette nature de document;E;O;O;O;'
      
        '3;?caption?;Choix impossible. Ce tiers ferm'#233' ne peut pas '#234'tre ap' +
        'pel'#233' pour cette nature de document;E;O;O;O;'
      
        '4;?caption?;Choix impossible. La nature du tiers associ'#233' '#224' l'#39'aff' +
        'aire n'#39'est pas compatible avec la nature de pi'#232'ce;E;O;O;O;'
      '5;?caption?;Vous devez renseigner un tiers valide;E;O;O;O;'
      '6;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;N;N;'
      
        '7;?caption?;Voulez-vous affecter ce d'#233'p'#244't sur toutes les lignes ' +
        'concern'#233'es ?;Q;YN;Y;N;'
      
        '8;?caption?;Choix impossible. Ce code tiers est incorrect;E;O;O;' +
        'O;'
      
        '9;?caption?;Cette pi'#232'ce ne peut pas '#234'tre modifi'#233'e ni transform'#233'e' +
        ', seule la consultation est autoris'#233'e;E;O;O;O;'
      
        '10;?caption?;Cet article est en rupture, confirmez-vous malgr'#233' t' +
        'out la quantit'#233' ?;Q;YN;Y;N;'
      '11;?caption?;Cet article n'#39'est pas disponible ;E;O;O;O;'
      '12;?caption?;Cet article n'#39'est plus g'#233'r'#233' ;E;O;O;O;'
      
        '13;?caption?;Le plafond de cr'#233'dit accord'#233' au tiers est d'#233'pass'#233' !' +
        ' ;E;O;O;O;'
      '14;?caption?;Le cr'#233'dit accord'#233' au tiers est d'#233'pass'#233' ! ;E;O;O;O;'
      
        '15;?caption?;Voulez-vous saisir ce taux dans la table de chancel' +
        'lerie ?;Q;YN;Y;N;O;'
      
        '16;?caption?;ATTENTION : Le taux en cours est de 1. Voulez-vous ' +
        'saisir ce taux dans la table de chancellerie;Q;YN;Y;N;O;'
      
        '17;?caption?;ATTENTION : la marge de la ligne est inf'#233'rieure au ' +
        'minimum requis;E;O;O;O;'
      
        '18;?caption?;La date que vous avez renseign'#233'e n'#39'est pas valide;E' +
        ';O;O;O;'
      
        '19;?caption?;La date que vous avez renseign'#233'e n'#39'est pas dans un ' +
        'exercice ouvert;E;O;O;O;'
      
        '20;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture;E;O;O;O;'
      
        '21;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture;E;O;O;O;'
      
        '22;?caption?;La date que vous avez renseign'#233'e est en dehors des ' +
        'limites autoris'#233'es;E;O;O;O;'
      
        '23;?caption?;ATTENTION : la marge de la ligne est inf'#233'rieure au ' +
        'minimum requis. Voulez-vous changer de code utilisateur ?;Q;YN;N' +
        ';N;'
      
        '24;?caption?;Le plafond de cr'#233'dit accord'#233' au tiers est d'#233'pass'#233'. ' +
        'Voulez-vous le faire valider par un superviseur ? ;Q;YN;N;N;'
      '25;?caption?;Vous ne pouvez pas saisir avant le ;E;O;O;O;'
      
        '26;?caption?;Cette pi'#232'ce contient d'#233'j'#224' des lignes. Reprise des l' +
        'ignes de l'#39'affaire impossible ;E;O;O;O;'
      
        '27;?caption?;Attention : Vous n'#39'avez pas affect'#233' d'#39'affaire sur c' +
        'ette pi'#232'ce ;E;O;O;O;'
      
        '28;?caption?;La date est ant'#233'rieure '#224' celle de derni'#232're cl'#244'ture ' +
        'de stock;E;O;O;O;'
      
        '29;?caption?;Confirmez-vous la suppression de la pi'#232'ce ?;Q;YN;N;' +
        'N;'
      
        '30;?caption?;Voulez-vous r'#233'percuter la date de livraison de l'#39'en' +
        't'#234'te sur toutes les lignes?;Q;YN;Y;N;O;'
      
        '31;?caption?;Voulez-vous r'#233'percuter la date de livraison de l'#39'en' +
        't'#234'te sur les lignes s'#233'lectionn'#233'es ?;Q;YN;Y;N;O;'
      '32;?caption?;Vous devez renseigner une devise;E;O;O;O;'
      
        '33;?caption?;Vous avez enregistr'#233' des acomptes ou des r'#232'glements' +
        '. Voulez-vous les supprimer ?;Q;YN;Y;N;O;'
      
        '34;?caption?;Ce client est en rouge ! Vous ne pouvez pas lui cr'#233 +
        'er de pi'#232'ce '#224'  ce niveau de risque;E;O;O;O;'
      
        '35;?caption?;Confirmez-vous la suppression de tout le paragraphe' +
        ' ?;Q;YN;N;N;'
      
        '36;?caption?;Cet article est tenu en stock, vous ne pouvez le co' +
        'ntremarquer;E;O;O;O;'
      
        '37;?caption?;Cette pi'#232'ce contient d'#233'j'#224' des lignes, vous ne pouve' +
        'z pas changer le fournisseur;E;O;O;O;'
      
        '38;?caption?;Voulez-vous affecter cette TVA sur toutes les ligne' +
        's concern'#233'es ?;Q;YN;Y;N;'
      
        '39;?caption?;Voulez-vous affecter ce r'#233'gime sur toute la pi'#232'ce ?' +
        ';Q;YN;Y;N;'
      '40;?caption?;Code inexistant;E;O;O;O;'
      '41;?caption?;Vous devez renseigner un motif par d'#233'faut;E;O;O;O;'
      
        '42;?caption?;Vous ne pouvez pas saisir une quantit'#233' n'#233'gative;E;O' +
        ';O;O;'
      
        '43;?caption?;Cette pi'#232'ce ne peut pas '#234'tre modifi'#233'e, seule la der' +
        'ni'#232're situation est modifiable;E;O;O;O;'
      '44;?caption?;Certains articles n'#39'ont pu '#234'tre recup'#233'r'#233's;E;O;O;O;'
      
        '45;?caption?;G'#233'n'#233'ration impossible : le param'#233'trage est incomple' +
        't;E;O;O;O;'
      '46;?caption?;Date inf'#233'rieure '#224' arr'#234't'#233' de p'#233'riode;E;O;O;O;'
      
        '47;?caption?;Saisie impossible. Cet article n'#39'est g'#233'rable qu'#39'en ' +
        'contremarque.;E;O;O;O;'
      
        '48;?caption?;La saisie par code barre n'#39'a pas '#233't'#233' valid'#233'e, voule' +
        'z-vous la valider ?;Q;YN;Y;N;'
      '49;?caption?;Le taux d'#39'escompte ne peut exc'#232'der 100%. ;E;O;O;O;'
      '50;?caption?;Le taux de remise ne peut exc'#232'der 100%. ;E;O;O;O;'
      
        '51;?caption?;Attention ! La quantit'#233' saisie est inf'#233'rieure au to' +
        'tal des quantit'#233's des pi'#232'ces suivantes.;E;O;O;O;'
      
        '52;?caption?;Vous ne pouvez pas supprimer cette ligne, d'#39'autres ' +
        'lignes y sont associ'#233'es.;E;O;O;O;'
      
        '53;?caption?;Vous ne pouvez pas changer l'#39'article de cette ligne' +
        '.;E;O;O;O;'
      
        '54;?caption?;Toutes les lignes de la pi'#232'ce d'#39'origine ont '#233't'#233' tra' +
        'nsform'#233'es.;E;O;O;O;'
      
        '55;?caption?;ATTENTION: La suppression est impossible car de l'#39'a' +
        'ctivit'#233' a '#233't'#233' saisie ;E;O;O;O; '
      
        '56;?caption?;La quantit'#233' de la ligne est sup'#233'rieure au maximum a' +
        'utoris'#233';E;O;O;O;'
      '57;?caption?;Le d'#233'p'#244't est obligatoire.;E;O;O;O;'
      '58;?caption?;Confirmez-vous le solde de cette ligne ?;Q;YN;N;N;'
      '59;?caption?;Voulez-vous d'#233'-solder de cette ligne ?;Q;YN;N;N;'
      
        '60;?caption?;Vous ne pouvez pas supprimer cette ligne, la ligne ' +
        'de la pi'#232'ce pr'#233'c'#233'dente est sold'#233'e.;E;O;O;O;'
      
        '61;?caption?;Confirmez-vous le solde de toutes les lignes de la ' +
        'piece ?;Q;YN;N;N;'
      
        '62;?caption?;La date que vous avez renseign'#233'e est hors des dates' +
        ' limites de saisies;E;O;O;O;'
      '')
    Left = 533
    Top = 167
  end
  object POPY: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 44
    Top = 301
    object POPYTS: TMenuItem
      Caption = 'Avenants / TS'
      OnClick = POPYTSClick
    end
    object SEPNN: TMenuItem
      Caption = '-'
    end
    object RepriseAvancPreGlob: TMenuItem
      Caption = 'Reprise avancement situation pr'#233'c'#233'dente'
      Visible = False
      OnClick = RepriseAvancPreGlobClick
    end
    object RecalcAvanc: TMenuItem
      Caption = 'Reprise avancement (Ligne)'
      Visible = False
      OnClick = RecalcAvancClick
    end
    object Sep11: TMenuItem
      Caption = '-'
      Visible = False
    end
    object VoirFranc: TMenuItem
      Caption = 'Voir la pi'#232'ce en franc'
      Visible = False
      OnClick = VoirFrancClick
    end
    object SepTva1000: TMenuItem
      Caption = '-'
      Visible = False
    end
    object MBREPARTTVA: TMenuItem
      Caption = 'R'#233'partition de TVA au 1/1000e'
      OnClick = MBREPARTTVAClick
    end
    object NSEPARATEUR: TMenuItem
      Caption = '-'
      Visible = False
    end
    object CpltEntete: TMenuItem
      Caption = 'Compl'#233'ment &Ent'#234'te'
      OnClick = CpltEnteteClick
    end
    object Librepiece: TMenuItem
      Caption = 'Zones libres pi'#232'ces'
      OnClick = LibrepieceClick
    end
    object VariablesMetres: TMenuItem
      Caption = '&Variables M'#233'tr'#233's'
      object VarApplication: TMenuItem
        Caption = '&Applications'
        OnClick = VariablesApplicationClick
      end
      object VarGenerale: TMenuItem
        Caption = '&G'#233'n'#233'rales'
        OnClick = VariablesGeneraleClick
      end
      object VarDocument: TMenuItem
        Caption = '&Document'
        OnClick = VariablesDocumentClick
      end
      object VarLigne: TMenuItem
        Caption = '&Ligne'
        OnClick = VariablesLigneClick
      end
    end
    object CpltLigne: TMenuItem
      Caption = 'Compl'#233'ment &Ligne'
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
    object MBTarifVisuOrigine: TMenuItem
      Caption = 'Justification &Tarifaire'
      OnClick = MBTarifVisuOrigineClick
    end
    object N9: TMenuItem
      Caption = '-'
    end
    object MBtarif: TMenuItem
      Caption = 'Mise '#224' jour tarif : ligne '#224' ligne'
      OnClick = MBtarifClick
    end
    object MBTarifGroupe: TMenuItem
      Caption = 'Mise '#224' jour tarif : group'#233'e pi'#232'ce'
      OnClick = MBTarifGroupeClick
    end
    object MBDatesLivr: TMenuItem
      Caption = 'Mise '#224' jour dates livraison lignes'
      OnClick = MBDatesLivrClick
    end
    object MBSoldeReliquat: TMenuItem
      Caption = 'Solder / Activer le reliquat'
      OnClick = MBSoldeReliquatClick
    end
    object MBSoldeTousReliquat: TMenuItem
      Caption = 'Solder / Activer tous les reliquats'
      OnClick = MBSoldeTousReliquatClick
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object MBDetailNomen: TMenuItem
      Caption = 'D'#233'tail nomenclature'
      OnClick = MBDetailNomenClick
    end
    object MBDetailLot: TMenuItem
      Caption = 'D'#233'tail lots/s'#233'ries'
      OnClick = MBDetailLotClick
    end
    object MBCtrlFact: TMenuItem
      Caption = 'Contr'#244'le des pi'#232'ces d'#39'achat'
      OnClick = MBCtrlFactClick
    end
    object AnalyseFiche1: TMenuItem
      Caption = 'Analyse'
      OnClick = AnalyseFiche1Click
    end
    object SIMULATIONRENTABILIT1: TMenuItem
      Caption = 'Simulation Rentabilit'#233
      OnClick = SIMULATIONRENTABILIT1Click
    end
    object MBModeVisu: TMenuItem
      Caption = 'Mode d'#39'affichage des d'#233'tails d'#39'ouvrage'
      OnClick = MBModeVisuClick
    end
    object Modebordereau: TMenuItem
      Caption = 'Mode P.V bloqu'#233
      ShortCut = 24689
      Visible = False
      OnClick = ModebordereauClick
    end
    object MnVisaPiece: TMenuItem
      Caption = 'Viser la pi'#232'ce'
      OnClick = MnVisaPieceClick
    end
    object VoirBPXSpigao: TMenuItem
      Caption = 'SPIGAO : Voir Bordereau'
      OnClick = VoirBPXSpigaoClick
    end
    object N4: TMenuItem
      Caption = '-'
    end
    object FraisDetail: TMenuItem
      Caption = 'Saisie des frais d'#233'taill'#233's de chantier '
      Visible = False
      OnClick = FraisDetailClick
    end
    object N8: TMenuItem
      Caption = '-'
      Visible = False
    end
    object BCOTRAITTABLEAU: TMenuItem
      Caption = 'Tableau r'#233'capitulatif des Intervenants'
    end
    object PAIESTPOC: TMenuItem
      Caption = 'Paiement Sous-traitants'
      Visible = False
    end
    object BPOPYCOTRAITSEP: TMenuItem
      Caption = '-'
    end
    object PlanningEntete: TMenuItem
      Caption = 'Planning Ent'#234'te'
    end
    object PlanningLigne: TMenuItem
      Caption = 'Planning Ligne '
    end
    object SEPY5: TMenuItem
      Caption = '-'
    end
    object CTRLOUV: TMenuItem
      Caption = 'Contr'#244'le conformit'#233' ouvrages'
      OnClick = CTRLOUVClick
    end
    object N5: TMenuItem
      Caption = '-'
    end
    object MBSGED: TMenuItem
      Caption = 'S.G.E.D.'
      object SGEDPiece: TMenuItem
        Caption = 'Pi'#232'ce'
        OnClick = SGEDPieceClick
      end
      object SGEDLigne: TMenuItem
        Caption = 'Ligne'
        OnClick = SGEDLigneClick
      end
    end
  end
  object PopD: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 72
    Top = 147
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
  object PopL: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 40
    Top = 442
    object MBAJUSTSIT: TMenuItem
      Caption = 'Ajustement situation'
      OnClick = MBAJUSTSITClick
    end
    object SEPAJUSTSIT: TMenuItem
      Caption = '-'
    end
    object TInsParag: TMenuItem
      Caption = 'Ins'#233'rer &Paragraphe'
      OnClick = TInsParagClick
    end
    object N6: TMenuItem
      Caption = '-'
    end
    object InsTextes: TMenuItem
      Caption = 'Ins'#233'rer texte memoris'#233
      Visible = False
      OnClick = InsTextesClick
    end
    object TInsLigne: TMenuItem
      Caption = 'Ins'#233'rer ligne'
      OnClick = TInsLigneClick
    end
    object TSupLigne: TMenuItem
      Caption = 'Supprimer ligne'
      OnClick = TSupLigneClick
    end
    object TFusionner: TMenuItem
      Caption = 'Fusionner les lignes'
      Visible = False
      OnClick = TFusionnerClick
    end
    object N7: TMenuItem
      Caption = '-'
    end
    object TCommentEnt: TMenuItem
      OnClick = TCommentEntClick
    end
    object TCommentPied: TMenuItem
      Caption = 'Commentaire en pied'
      OnClick = TCommentPiedClick
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
      
        '1;?caption?;Vous ne pouvez pas enregistrer une pi'#232'ce vide;E;O;O;' +
        'O;'
      
        '2;?caption?;Vous ne pouvez pas enregistrer une pi'#232'ce sans articl' +
        'es;E;O;O;O;'
      
        '3;?caption?;Vous ne pouvez pas enregistrer une pi'#232'ce sans mouvem' +
        'ents;E;O;O;O;'
      
        '4;?caption?;Vous ne pouvez pas enregistrer une pi'#232'ce '#224' cette dat' +
        'e;E;O;O;O;'
      
        '5;?caption?;Enregistrement impossible : l'#39'acompte est sup'#233'rieur ' +
        'au total de la pi'#232'ce;E;O;O;O;'
      
        '6;?caption?;Enregistrement impossible : la devise est incorrecte' +
        ';E;O;O;O;'
      '7;?caption?;La pi'#232'ce d'#39'origine est d'#233'j'#224' trait'#233'e;E;O;O;O;'
      '8;?caption?;Le contr'#244'le facture n'#39'est pas valide;E;O;O;O;'
      '9;?caption?;Le circuit est obligatoire;E;O;O;O;'
      
        '10;?caption?;Enregistrement impossible : des montants sont '#224' z'#233'r' +
        'o;E;O;O;O;'
      
        '11;?caption?;Enregistrement impossible : Vous devez renseigner l' +
        'e nombre d'#39'heures;E;O;O;O;')
    Left = 640
    Top = 231
  end
  object HAveugle: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ATTENTION : code barre inconnu.'
      
        '1;?Caption?;La quantit'#233' de cet article est sup'#233'rieure '#224' celle at' +
        'tendue dans le document. #13 Voulez-vous l'#39'ajouter au document ?' +
        ';Q;YN;Y;N;'
      'ATTENTION : vous ne pouvez pas saisir une quantit'#233' n'#233'gative.'
      
        '3;?Caption?;Voulez-vous supprimer tous les articles saisis ?;Q;Y' +
        'N;Y;N;'
      'ATTENTION : cet article ne correspond pas '#224' ce fournisseur.'
      '')
    Left = 557
    Top = 226
  end
  object PopV: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 40
    Top = 147
    object VPiece: TMenuItem
      Caption = 'Pi'#232'ce'
      OnClick = VPieceClick
    end
    object VLigne: TMenuItem
      Caption = 'Ligne'
      OnClick = VLigneClick
    end
    object VSepare: TMenuItem
      Caption = '-'
    end
    object SPiece: TMenuItem
      Caption = 'Stock ent'#234'te'
      OnClick = SPieceClick
    end
    object SLigne: TMenuItem
      Caption = 'Stock ligne'
      OnClick = SLigneClick
    end
  end
  object PopA: TPopupMenu
    AutoPopup = False
    Left = 41
    Top = 390
    object Pourlaligne: TMenuItem
      Caption = 'Pour la ligne'
      OnClick = PourlaligneClick
    end
    object Pourledocument: TMenuItem
      Caption = 'Pour le document'
      OnClick = PourledocumentClick
    end
  end
  object PopOuvrage: TPopupMenu
    AutoPopup = False
    Left = 40
    Top = 204
    object DetailOuvrage: TMenuItem
      Caption = 'D'#233'tail de l'#39'ouvrage'
      OnClick = DetailOuvrageClick
    end
    object SousDetailAff: TMenuItem
      Caption = 'Affichage/Retrait des sous d'#233'tails'
      OnClick = SousDetailAffClick
    end
  end
  object ImTypeArticle: THImageList
    GlobalIndexImages.Strings = (
      'Z0054_S16G1'
      'Z0153_S16G1'
      'Z0507_S16G1'
      'Z0146_S16G1'
      'Z0160_S16G1'
      'Z0150_S16G1'
      'Z0415_S16G1'
      'Z0885_S16G1'
      'Z1439_S16G1'
      'Z0284_S16G1')
    Height = 18
    Width = 18
    Left = 651
    Top = 152
  end
  object ODExcelFile: TOpenDialog
    DefaultExt = '*.xlsx'
    Filter = 
      'Fichier Excel 2007|*.xlsx|Fichier Excel 97-2003|*.xls|Tous|*.xls' +
      ';*.xlsx'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Left = 241
    Top = 155
  end
  object ImgListPlusMoins: THImageList
    GlobalIndexImages.Strings = (
      'Z0533_S16G1'
      'Z0620_S16G1')
    Left = 688
    Top = 152
  end
  object PopInsertLig: TPopupMenu
    AutoPopup = False
    Left = 40
    Top = 252
    object MnInsSousDet: TMenuItem
      Caption = 'Insertion d'#39'un sous d'#233'tail de l'#39'ouvrage'
      OnClick = MnInsSousDetClick
    end
    object MnInsLigStd: TMenuItem
      Caption = 'Insertion d'#39'une ligne dans le document'
      OnClick = MnInsLigStdClick
    end
  end
  object POPDEMPRIX: TPopupMenu
    OwnerDraw = True
    Left = 700
    Top = 364
    object MnVisuDemPrix: TMenuItem
      Caption = 'Gestion des Demandes De Prix'
      OnClick = MnVisuDemPrixClick
    end
    object mniN11: TMenuItem
      Caption = '-'
    end
    object MnConstitueDemPrix: TMenuItem
      Caption = 'Affectation/R'#233'actualisation des Articles'
      OnClick = MnConstitueDemPrixClick
    end
    object mniN10: TMenuItem
      Caption = '-'
    end
    object MnMajDemPrix: TMenuItem
      Caption = 'Mettre '#224' Jour les Demandes de Prix'
      OnClick = MnMajDemPrixClick
    end
  end
  object PopIntExcel: TPopupMenu
    Left = 40
    Top = 477
    object IntgrationdunfichierExcel1: TMenuItem
      Caption = 'Int'#233'gration d'#39'un fichier Excel'
      OnClick = BintegreExcelClick
    end
    object N10: TMenuItem
      Caption = '-'
    end
    object Paramtragedelimport1: TMenuItem
      Caption = 'Param'#233'trage de l'#39'import'
      OnClick = ParamImportExcel
    end
  end
  object PopNumauto: THPopupMenu
    AutoPopup = False
    Left = 856
    Top = 136
    object Pnum: TMenuItem
      Caption = 'Num'#233'rotation'
      OnClick = BRENUMClick
    end
    object DelNumAuto: TMenuItem
      Caption = 'Suppression de la num'#233'rotation'
      OnClick = DelNumAutoClick
    end
  end
  object HimgTOT: THImageList
    GlobalIndexImages.Strings = (
      'Z0415_S16G1'
      'Z0397_S16G1')
    DrawingStyle = dsTransparent
    Height = 18
    Width = 18
    Left = 725
    Top = 150
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 578
    Top = 168
  end
  object POPGEDBSV: TPopupMenu
    OwnerDraw = True
    Left = 664
    Top = 364
    object MnBSVSTOCKE: TMenuItem
      Caption = 'Stockage du document'
      OnClick = MnBSVSTOCKEClick
    end
    object MnBSVVISU: TMenuItem
      Caption = 'Visualisation du document'
      OnClick = MnBSVVISUClick
    end
  end
end

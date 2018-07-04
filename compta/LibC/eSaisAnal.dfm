object FeSaisAnal: TFeSaisAnal
  Left = 353
  Top = 230
  Width = 467
  Height = 328
  HelpContext = 7244100
  ActiveControl = GSA
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = '*'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 451
    Height = 9
  end
  object DockRight: TDock97
    Left = 442
    Top = 9
    Width = 9
    Height = 252
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 252
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 261
    Width = 451
    Height = 28
    Position = dpBottom
    object Outils: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 0
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Calcul du solde'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeClick
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object BVentilType: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Appel derni'#232're ventilation type'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BVentilTypeClick
        GlobalIndexImage = 'Z0566_S16G1'
        IsControl = True
      end
      object BComplement: TToolbarButton97
        Tag = 1
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Informations compl'#233'mentaires'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BComplementClick
        GlobalIndexImage = 'Z0105_S16G2'
        IsControl = True
      end
      object BZoom: TToolbarButton97
        Tag = 1
        Left = 81
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Zoom Section'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BZoomClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BCALCULQTE: TToolbarButton97
        Tag = 1
        Left = 108
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Recalcul des quantit'#233's de l'#39'axe courant'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BCALCULQTEClick
        GlobalIndexImage = 'Z0123_S16G1'
        IsControl = True
      end
    end
    object Valide97: TToolbar97
      Left = 361
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 361
      TabOrder = 1
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Valider la saisie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Abandonner'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
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
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
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
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 433
    Height = 252
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object PEntete: TPanel
      Left = 0
      Top = 0
      Width = 433
      Height = 45
      Align = alTop
      TabOrder = 0
      object H_GENERAL: THLabel
        Left = 8
        Top = 8
        Width = 45
        Height = 13
        AutoSize = False
        Caption = 'Compte'
      end
      object G_GENERAL: THLabel
        Left = 68
        Top = 8
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'G_GENERAL'
      end
      object H_REFINTERNE: THLabel
        Left = 8
        Top = 27
        Width = 53
        Height = 13
        AutoSize = False
        Caption = 'R'#233'f/Lib'
      end
      object E_REFINTERNE: THLabel
        Left = 68
        Top = 27
        Width = 121
        Height = 13
        AutoSize = False
        Caption = 'E_REFINTERNE'
      end
      object E_LIBELLE: THLabel
        Left = 200
        Top = 27
        Width = 217
        Height = 13
        AutoSize = False
        Caption = 'E_LIBELLE'
      end
      object H_MONTANTECR: THLabel
        Left = 200
        Top = 8
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'Montant '#224' r'#233'partir'
      end
      object ISigneEuro: TImage
        Left = 426
        Top = 7
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
      object E_MONTANTECR: THNumEdit
        Tag = 1
        Left = 300
        Top = 6
        Width = 121
        Height = 20
        TabStop = False
        AutoSize = False
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        OnEnter = E_MontantEcrEnter
        OnExit = E_MONTANTECRExit
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object Cache: THCpteEdit
        Left = 144
        Top = 6
        Width = 17
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        Text = '!!!'
        Visible = False
        ZoomTable = tzSection
        Vide = False
        Bourre = True
        okLocate = False
        SynJoker = False
      end
    end
    object PPied: TPanel
      Left = 0
      Top = 195
      Width = 433
      Height = 57
      Align = alBottom
      BorderStyle = bsSingle
      TabOrder = 1
      object H_SOLDE: THLabel
        Left = 4
        Top = 31
        Width = 49
        Height = 13
        AutoSize = False
        Caption = 'Solde'
      end
      object H_TOTAL: THLabel
        Left = 212
        Top = 7
        Width = 53
        Height = 13
        AutoSize = False
        Caption = 'Total saisi'
      end
      object H_RESTE: THLabel
        Left = 212
        Top = 31
        Width = 49
        Height = 13
        AutoSize = False
        Caption = 'Reste'
      end
      object TFType: THLabel
        Left = 4
        Top = 8
        Width = 53
        Height = 13
        Caption = '&Ventil.Type'
        FocusControl = CVType
      end
      object BSauveVentil: TToolbarButton97
        Left = 184
        Top = 4
        Width = 21
        Height = 21
        Hint = 'Enregistrer comme ventilation type sur l'#39'axe en cours'
        ParentShowHint = False
        ShowHint = True
        OnClick = BSauveVentilClick
        GlobalIndexImage = 'Z0393_S16G1'
      end
      object S_SOLDE: THNumEdit
        Tag = 2
        Left = 60
        Top = 28
        Width = 121
        Height = 21
        Color = clBtnFace
        Enabled = False
        TabOrder = 0
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object E_TOTAL: THNumEdit
        Tag = 1
        Left = 268
        Top = 4
        Width = 113
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object E_RESTE: THNumEdit
        Tag = 1
        Left = 268
        Top = 28
        Width = 113
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object E_TOTPOURC: THNumEdit
        Left = 392
        Top = 4
        Width = 61
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntPercentage
        UseRounding = True
        Validate = False
      end
      object E_RESTEPOURC: THNumEdit
        Left = 392
        Top = 28
        Width = 61
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 4
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntPercentage
        UseRounding = True
        Validate = False
      end
      object CVType: THValComboBox
        Left = 60
        Top = 4
        Width = 121
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = CVTypeChange
        TagDispatch = 0
        DataType = 'TTVENTILTYPE'
      end
    end
    object PGA: TPanel
      Left = 0
      Top = 45
      Width = 433
      Height = 150
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 2
      object GSA: THGrid
        Left = 0
        Top = 25
        Width = 433
        Height = 125
        Align = alClient
        BorderStyle = bsNone
        ColCount = 6
        DefaultRowHeight = 18
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
        TabOrder = 0
        OnDblClick = GSADblClick
        OnKeyPress = GSAKeyPress
        OnMouseDown = GSAMouseDown
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = GSARowEnter
        OnRowExit = GSARowExit
        OnCellEnter = GSACellEnter
        OnCellExit = GSACellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          22
          66
          165
          55
          88
          64)
        RowHeights = (
          18
          18)
      end
      object TA: TTabControl
        Left = 0
        Top = 0
        Width = 433
        Height = 25
        Align = alTop
        TabOrder = 1
        Tabs.Strings = (
          'Axe n'#176' 1'
          'Axe n'#176' 2'
          'Axe n'#176' 3'
          'Axe n'#176' 4'
          'Axe n'#176' 5')
        TabIndex = 0
        OnChange = TAChange
        OnChanging = TAChanging
      end
      object BInsert: THBitBtn
        Tag = 1000
        Left = 100
        Top = 47
        Width = 28
        Height = 25
        Hint = 'Nouvelle ligne'
        Caption = 'Ins.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        Visible = False
        OnClick = BInsertClick
        Margin = 2
      end
      object BSDel: THBitBtn
        Tag = 1000
        Left = 132
        Top = 47
        Width = 28
        Height = 25
        Hint = 'Supprimer ligne'
        Caption = 'Del.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        Visible = False
        OnClick = BSDelClick
        Margin = 2
      end
    end
    object PVentil: TPanel
      Left = 76
      Top = 112
      Width = 321
      Height = 65
      TabOrder = 3
      Visible = False
      object H_CODEVENTIL: THLabel
        Left = 8
        Top = 40
        Width = 25
        Height = 13
        Caption = 'Code'
      end
      object HLabel1: THLabel
        Left = 92
        Top = 40
        Width = 30
        Height = 13
        Caption = 'Libell'#233
      end
      object H_TITREVENTIl: TLabel
        Left = 28
        Top = 12
        Width = 273
        Height = 16
        Alignment = taCenter
        Caption = 'Cr'#233'ation d'#39'une nouvelle ventilation type'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Y_CODEVENTIL: TEdit
        Left = 40
        Top = 36
        Width = 37
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 3
        TabOrder = 0
      end
      object Y_LIBELLEVENTIL: TEdit
        Left = 128
        Top = 36
        Width = 149
        Height = 21
        MaxLength = 35
        TabOrder = 1
      end
      object BNewVentil: THBitBtn
        Tag = 1
        Left = 285
        Top = 33
        Width = 28
        Height = 27
        Hint = 'Valider la cr'#233'ation'
        TabOrder = 2
        OnClick = BNewVentilClick
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object PFenCouverture: TPanel
        Left = 1
        Top = 1
        Width = 319
        Height = 9
        Align = alTop
        BevelInner = bvLowered
        BevelOuter = bvNone
        Color = clTeal
        Enabled = False
        TabOrder = 3
      end
    end
  end
  object HMessLigne: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Ventilations analytiques;Votre saisie est incorrecte : section' +
        ' analytique absente ou incorrecte;W;O;O;O;'
      
        '1;Ventilations analytiques;Vous devez saisir un montant.;W;O;O;O' +
        ';'
      
        '2;Ventilations analytiques;Attention : le montant en devise pivo' +
        't est nul;E;O;O;O;'
      
        '3;Ventilations analytiques;Vous ne pouvez pas saisir un montant ' +
        'n'#233'gatif.;W;O;O;O;'
      
        '4;Ventilations analytiques;La section que vous avez renseign'#233'e e' +
        'st ferm'#233'e.;E;O;O;O;'
      
        '5;?caption?;Attention : cette section est ferm'#233'e. Vous ne pouvez' +
        ' plus l'#39'utiliser en saisie;E;O;O;O;')
    Left = 37
    Top = 123
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 416
    Top = 159
  end
  object HMessPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Ventilations analytiques;La somme des ventilations doit '#234'tre '#233 +
        'gale '#224' 100%;W;O;O;O;'
      
        '1;Ventilations analytiques;Vous devez saisir au moins une ligne;' +
        'W;O;O;O;'
      '2;LIBRE;LIBRE;E;O;O;O;'
      
        '3;Ventilations analytiques;La ventilation en devise pivot n'#39#233'qui' +
        'libre pas le montant;W;O;O;O;'
      
        '4;Ventilations analytiques;Vous devez saisir un code et un libel' +
        'l'#233';W;O;O;O;'
      
        '5;Ventilations analytiques;Le code ou le libell'#233' existe d'#233'j'#224'.;W;' +
        'O;O;O;'
      
        '6;Ventilations analytiques;Certaines sections sur le compte g'#233'n'#233 +
        'ral ne sont pas compatibles avec la d'#233'finition du budget. Confir' +
        'mez-vous la validation ?;Q;YN;N;N;')
    Left = 36
    Top = 167
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 416
    Top = 112
  end
end

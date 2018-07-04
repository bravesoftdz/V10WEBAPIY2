object FDelVisuE: TFDelVisuE
  Left = 401
  Top = 263
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSingle
  Caption = 'Liste des pi'#232'ces'
  ClientHeight = 266
  ClientWidth = 564
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HP6: TPanel
    Left = 0
    Top = 231
    Width = 564
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object H_NB: TLabel
      Left = 12
      Top = 12
      Width = 231
      Height = 13
      Caption = 'Nombre de pi'#232'ces g'#233'n'#233'r'#233'es en '#233'criture courante'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object BAnnuler: THBitBtn
      Left = 500
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 532
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object BImprimer: THBitBtn
      Left = 468
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BZoomPiece: THBitBtn
      Tag = 1
      Left = 436
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Zoom de l'#39#233'criture'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BZoomPieceClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0016_S16G1'
      IsControl = True
    end
  end
  object GDEL: THGrid
    Tag = 1
    Left = 0
    Top = 0
    Width = 564
    Height = 231
    Align = alClient
    ColCount = 25
    Ctl3D = True
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 15
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 1
    OnDblClick = GDELDblClick
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = -1
    SortEnabled = False
    SortRowExclude = -1
    TwoColors = True
    AlternateColor = clSilver
    ColWidths = (
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64)
  end
  object HPieces: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Liste des '#233'critures d'#233'truites'
      'Liste des '#233'critures g'#233'n'#233'r'#233'es'
      'Liste des '#233'critures non g'#233'n'#233'r'#233'es'
      'Liste des '#233'critures g'#233'n'#233'r'#233'es par la r'#233'gularisation de lettrage'
      
        'Liste des '#233'critures g'#233'n'#233'r'#233'es par la r'#233'gularisation d'#39#233'cart de ch' +
        'ange'
      
        'Liste des '#233'critures analytiques g'#233'n'#233'r'#233'es par le transfert inter-' +
        'sections'
      'Compte-rendu du lettrage automatique'
      'Liste des '#233'critures cr'#233#233'es par la g'#233'n'#233'ration des abonnements'
      'Liste des '#233'critures de r'#233'-imputations'
      'Liste des '#233'critures de budget d'#233'truites'
      'Liste des '#233'critures cr'#233#233'es par la g'#233'n'#233'ration d'#39'extournes'
      'Liste des '#233'critures export'#233'es'
      'Liste des '#233'critures analytiques g'#233'n'#233'r'#233'es par la r'#233'-imputation'
      
        'Liste des '#233'critures analytiques g'#233'n'#233'r'#233'es par r'#233'partition seconda' +
        'ire'
      
        'Liste des '#233'critures g'#233'n'#233'r'#233'es par la r'#233'gularisation d'#39#233'cart de co' +
        'nversion'
      'Liste des '#233'critures g'#233'n'#233'r'#233'es en lettrage manuel'
      'Liste des engagements li'#233's '#224' la facture selectionn'#233'e'
      'Liste des factures li'#233'es '#224' l'#39'engagement selectionn'#233
      'Liste des '#233'critures de Tiers payeurs'
      'Liste des '#233'critures d'#39'emprunt')
    Left = 15
    Top = 152
  end
  object HCaption: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Nombre d'#39#233'critures d'#233'truites : '
      'Nombre d'#39#233'critures g'#233'n'#233'r'#233'es en '#233'criture courante : '
      'Nombre d'#39#233'critures non g'#233'n'#233'r'#233'es : '
      'Nombre d'#39#233'critures de r'#233'gularisation : '
      'Nombre d'#39#233'critures d'#39#233'cart de change : '
      'Nombre d'#39#233'critures analytiques g'#233'n'#233'r'#233'es : '
      'Nombre de comptes trait'#233's : '
      'Nombre d'#39#233'critures d'#39'abonnement g'#233'n'#233'r'#233'es : '
      'Nombre d'#39#233'critures de r'#233'-imputation g'#233'n'#233'r'#233'es : '
      'Nombre d'#39#233'critures de budget d'#233'tuites : '
      'Nombre d'#39#233'critures d'#39'extourne g'#233'n'#233'r'#233'es : '
      'Nombre d'#39#233'critures export'#233'es : '
      'Nombre d'#39#233'critures analytiques g'#233'n'#233'r'#233'es : '
      'Nombre d'#39#233'critures analytiques g'#233'n'#233'r'#233'es : '
      'Nombre d'#39#233'critures d'#39#233'cart de conversion : '
      'Nombre d'#39#233'critures d'#39#233'critures g'#233'n'#233'r'#233'es :'
      'Nombre d'#39' engagements rapproch'#233's :'
      'Nombre de factures rapproch'#233'es'
      'Nombre d'#39#233'critures de Tiers payeurs g'#233'n'#233'r'#233'es'
      'Nombre d'#39#233'critures d'#39'emprunt g'#233'n'#233'r'#233'es')
    Left = 68
    Top = 152
  end
  object HButton: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Justificatif de solde')
    Left = 120
    Top = 152
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 172
    Top = 152
  end
end

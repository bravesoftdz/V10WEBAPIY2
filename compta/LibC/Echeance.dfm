object FEcheance: TFEcheance
  Left = 277
  Top = 173
  HelpContext = 7244200
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'R'#233'partition des '#233'ch'#233'ances'
  ClientHeight = 414
  ClientWidth = 664
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  Scaled = False
  ShowHint = True
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonPanel: TPanel
    Left = 0
    Top = 347
    Width = 664
    Height = 67
    Align = alBottom
    BevelOuter = bvNone
    Caption = ' '
    TabOrder = 0
    DesignSize = (
      664
      67)
    object Label18: TLabel
      Left = 44
      Top = 30
      Width = 45
      Height = 13
      Anchors = [akTop, akRight]
      Caption = 'D'#233'j'#224' saisi'
    end
    object Label19: TLabel
      Left = 44
      Top = 50
      Width = 43
      Height = 13
      Anchors = [akTop, akRight]
      Caption = 'Reste d'#251
    end
    object Label17: TLabel
      Left = 44
      Top = 10
      Width = 55
      Height = 13
      Anchors = [akTop, akRight]
      Caption = 'Net '#224' payer'
    end
    object ISigneEuro: TImage
      Left = 250
      Top = 13
      Width = 16
      Height = 16
      Anchors = [akTop, akRight]
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
    object BValide: TToolbarButton97
      Tag = 1
      Left = 245
      Top = 39
      Width = 27
      Height = 27
      Hint = 'Valider les '#233'ch'#233'ances'
      Anchors = [akTop, akRight]
      OnClick = BValideClick
      GlobalIndexImage = 'Z0127_S16G1'
    end
    object ToolbarButton971: TToolbarButton97
      Tag = 1
      Left = 276
      Top = 39
      Width = 27
      Height = 27
      Hint = 'Fermer'
      Anchors = [akTop, akRight]
      Cancel = True
      ModalResult = 2
      OnClick = BAbandonClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: TToolbarButton97
      Tag = 1
      Left = 307
      Top = 39
      Width = 27
      Height = 27
      Hint = 'Aide'
      Anchors = [akTop, akRight]
      Cancel = True
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object BSolde: TToolbarButton97
      Tag = 1
      Left = 276
      Top = 8
      Width = 27
      Height = 27
      Hint = 'Solder sur la ligne'
      Anchors = [akTop, akRight]
      NumGlyphs = 2
      OnClick = BSoldeClick
      GlobalIndexImage = 'Z0051_S16G2'
    end
    object BRepart: TToolbarButton97
      Tag = 1
      Left = 307
      Top = 8
      Width = 27
      Height = 27
      Hint = 'R'#233'partir le montant sur toutes les '#233'ch'#233'ances'
      Anchors = [akTop, akRight]
      NumGlyphs = 2
      OnClick = BRepartClick
      GlobalIndexImage = 'Z0041_S16G2'
    end
    object FSaisie: THNumEdit
      Left = 128
      Top = 24
      Width = 109
      Height = 21
      Anchors = [akTop, akRight]
      Color = clBtnFace
      Ctl3D = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object FReste: THNumEdit
      Left = 128
      Top = 45
      Width = 109
      Height = 21
      Anchors = [akTop, akRight]
      Color = clBtnFace
      Ctl3D = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object FAPayer: THNumEdit
      Left = 128
      Top = 3
      Width = 109
      Height = 21
      Anchors = [akTop, akRight]
      Color = clBtnFace
      Ctl3D = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = False
      Validate = False
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 0
    Width = 664
    Height = 57
    Align = alTop
    Caption = ' '
    TabOrder = 1
    object Label1: TLabel
      Left = 2
      Top = 8
      Width = 93
      Height = 13
      Caption = '&Condition r'#232'glement'
      FocusControl = FModeRegle
    end
    object Label2: TLabel
      Left = 2
      Top = 32
      Width = 53
      Height = 13
      Caption = 'Mode initial'
      FocusControl = FModeInit
    end
    object BZoom: TToolbarButton97
      Tag = 1
      Left = 284
      Top = 3
      Width = 26
      Height = 23
      Hint = 'Zoom conditions de r'#232'glement'
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
      OnClick = BZoomClick
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
    object FModeRegle: THValComboBox
      Left = 96
      Top = 4
      Width = 186
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FModeRegleChange
      TagDispatch = 0
      DataType = 'TTMODEREGLE'
    end
    object FModeInit: THValComboBox
      Left = 96
      Top = 28
      Width = 186
      Height = 21
      Style = csDropDownList
      Color = clBtnFace
      Ctl3D = False
      Enabled = False
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 1
      TagDispatch = 0
      DataType = 'TTMODEREGLE'
    end
    object FVerouiller: TCheckBox
      Left = 285
      Top = 29
      Width = 58
      Height = 17
      Caption = 'Bloquer'
      TabOrder = 2
      OnClick = FVerouillerClick
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 57
    Width = 664
    Height = 290
    Align = alClient
    DefaultColWidth = 20
    DefaultRowHeight = 18
    RowCount = 13
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goTabs]
    ScrollBars = ssNone
    TabOrder = 2
    OnExit = FListeExit
    SortedCol = -1
    Titres.Strings = (
      'N'#176';C;R;'
      'Mode;G;S;'
      'B'#233'n'#233'ficiaire;G;R;'
      'Montant;D;R;'
      'Ech'#233'ance;C;D;')
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = FListeRowEnter
    OnCellExit = FListeCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ColWidths = (
      20
      121
      336
      91
      86)
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;R'#233'partition des '#233'ch'#233'ances;Erreur de r'#233'partition : le montant n' +
        #39'est pas r'#233'parti compl'#232'tement.;W;O;O;O;'
      
        '1;R'#233'partition des '#233'ch'#233'ances;Le montant est trop '#233'lev'#233' pour ce mo' +
        'de de paiement.Voulez-vous passer au mode de remplacement ?;Q;YN' +
        ';Y;N;'
      
        '2;R'#233'partition des '#233'ch'#233'ances;Vous ne pouvez pas renseigner des mo' +
        'ntants n'#233'gatifs.;W;O;O;O;'
      
        '3;R'#233'partition des '#233'ch'#233'ances;Les dates d'#39#233'ch'#233'ance doivent respect' +
        'er la plage de saisie autoris'#233'e;W;O;O;O;')
    Left = 100
    Top = 140
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 56
    Top = 244
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 168
    Top = 208
  end
end

object FBudVent: TFBudVent
  Left = 250
  Top = 166
  HelpContext = 15217100
  ActiveControl = FListe
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSingle
  Caption = 'Ventilation budg'#233'taire'
  ClientHeight = 363
  ClientWidth = 382
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 382
    Height = 9
  end
  object DockRight: TDock97
    Left = 373
    Top = 9
    Width = 9
    Height = 326
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 326
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 335
    Width = 382
    Height = 28
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 288
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 288
      TabOrder = 0
      object BValide: TToolbarButton97
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
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Fermer'
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
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
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
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object Outils97: TToolbar97
      Left = 6
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 6
      TabOrder = 1
      object H_CVType: THLabel
        Left = 0
        Top = 5
        Width = 59
        Height = 13
        Caption = ' Ventil.Type '
        FocusControl = CVType
      end
      object BSauveVentil: TToolbarButton97
        Left = 182
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Enregistrer comme ventilation type du budget'
        ParentShowHint = False
        ShowHint = True
        OnClick = BSauveVentilClick
        GlobalIndexImage = 'Z0393_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 209
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Supprimer'
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 242
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Solder sur la section d'#39'attente'
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeClick
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object ToolbarSep971: TToolbarSep97
        Left = 236
        Top = 0
      end
      object CVType: THValComboBox
        Left = 59
        Top = 1
        Width = 123
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = CVTypeChange
        TagDispatch = 0
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 364
    Height = 326
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object FListe: THGrid
      Left = 0
      Top = 27
      Width = 364
      Height = 220
      Align = alClient
      ColCount = 3
      DefaultRowHeight = 16
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Comptes;G;'
        'Pourcentage;D;'
        'Montant;D;')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnCellEnter = FListeCellEnter
      OnCellExit = FListeCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        130
        84
        140)
    end
    object PTot: TPanel
      Left = 0
      Top = 247
      Width = 364
      Height = 79
      Align = alBottom
      BevelInner = bvLowered
      BevelOuter = bvNone
      Enabled = False
      TabOrder = 1
      object TTot: TLabel
        Left = 9
        Top = 58
        Width = 58
        Height = 13
        Caption = 'Total ventil'#233
      end
      object Label1: TLabel
        Left = 7
        Top = 27
        Width = 77
        Height = 13
        Caption = 'Reste '#224' ventiler '
      end
      object Label2: TLabel
        Left = 8
        Top = 40
        Width = 110
        Height = 13
        Caption = '(Hors compte d'#39'attente)'
      end
      object TTotIni: TLabel
        Left = 142
        Top = 9
        Width = 70
        Height = 13
        Alignment = taRightJustify
        Caption = 'Total '#224' ventiler'
      end
      object TotalP: THNumEdit
        Left = 134
        Top = 54
        Width = 80
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object TotalM: THNumEdit
        Left = 220
        Top = 54
        Width = 137
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
      object ResteP: THNumEdit
        Left = 134
        Top = 29
        Width = 80
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        TabOrder = 2
        UseRounding = True
        Validate = False
      end
      object ResteM: THNumEdit
        Left = 220
        Top = 29
        Width = 137
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        TabOrder = 3
        UseRounding = True
        Validate = False
      end
      object TotIni: THNumEdit
        Left = 220
        Top = 5
        Width = 137
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        TabOrder = 4
        UseRounding = True
        Validate = False
      end
    end
    object PVentil: TPanel
      Left = 6
      Top = 120
      Width = 353
      Height = 65
      TabOrder = 2
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
        Left = 40
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
        TabOrder = 0
      end
      object Y_LIBELLEVENTIL: TEdit
        Left = 128
        Top = 36
        Width = 149
        Height = 21
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
        Spacing = -1
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object PFenCouverture: TPanel
        Left = 1
        Top = 1
        Width = 351
        Height = 9
        Align = alTop
        BevelInner = bvLowered
        BevelOuter = bvNone
        Color = clTeal
        Enabled = False
        TabOrder = 3
      end
      object BAbandonVentil: THBitBtn
        Left = 318
        Top = 33
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnClick = BAbandonVentilClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
    end
    object Panel1: TPanel
      Left = 0
      Top = 0
      Width = 364
      Height = 27
      Align = alTop
      BevelInner = bvRaised
      BevelOuter = bvNone
      TabOrder = 3
      object TitreVentil: TLabel
        Left = 7
        Top = 5
        Width = 367
        Height = 19
        AutoSize = False
        Caption = 'Ventilations budg'#233'taires : '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Saisie des ventilations budg'#233'taires'
      'Modification des ventilations budg'#233'taires'
      'Consultation des ventilations budg'#233'taires'
      
        '3;Ventilation budg'#233'taire;Voulez-vous enregistrer les modificatio' +
        'ns ?;Q;YNC;Y;Y;'
      
        '4;Ventilation budg'#233'taire;La ventilation n'#39'est pas '#224' 100% : le mo' +
        'ntant total et les pourcentages seront recalcul'#233's. Confirmez-vou' +
        's l'#39'enregistrement ?;Q;YNC;Y;Y;'
      'Solder sur le compte budg'#233'taire d'#39'attente'
      
        '6;Ventilations analytiques;Vous devez saisir un code et un libel' +
        'l'#233';W;O;O;O;'
      
        '7;Ventilations analytiques;Le code ou le libell'#233' existe d'#233'j'#224'.;W;' +
        'O;O;O;'
      
        '8;Ventilation budg'#233'taire;Confirmez-vous la suppression de la ven' +
        'tilation type ?;Q;YNC;Y;Y;'
      ' ')
    Left = 76
    Top = 200
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 200
  end
end

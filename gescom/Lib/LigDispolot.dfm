object FLigDispolot: TFLigDispolot
  Left = 376
  Top = 239
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  ClientHeight = 317
  ClientWidth = 418
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GQL: THGrid
    Tag = 1
    Left = 0
    Top = 49
    Width = 418
    Height = 165
    Align = alClient
    ColCount = 3
    DefaultRowHeight = 18
    RowCount = 50
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
    ScrollBars = ssVertical
    TabOrder = 0
    OnDblClick = GQLDblClick
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = GQLRowEnter
    OnRowExit = GQLRowExit
    OnCellEnter = GQLCellEnter
    OnCellExit = GQLCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    OnElipsisClick = GQLElipsisClick
    RowHeights = (
      18
      18
      18
      17
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
  object PTETE: THPanel
    Left = 0
    Top = 0
    Width = 418
    Height = 49
    Align = alTop
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TARTICLE: THLabel
      Left = 13
      Top = 9
      Width = 29
      Height = 13
      Caption = 'Article'
    end
    object TDEPOT: THLabel
      Left = 13
      Top = 28
      Width = 29
      Height = 13
      Caption = 'D'#233'p'#244't'
    end
    object EArticle: THLabel
      Left = 50
      Top = 9
      Width = 36
      Height = 13
      Caption = 'EArticle'
    end
    object EDepot: THLabel
      Left = 50
      Top = 28
      Width = 36
      Height = 13
      Caption = 'EDepot'
    end
  end
  object PPIED: THPanel
    Left = 0
    Top = 234
    Width = 418
    Height = 52
    Align = alBottom
    FullRepaint = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    AutoResize = True
    object TNReste: THLabel
      Left = 13
      Top = 31
      Width = 28
      Height = 13
      Caption = 'Reste'
    end
    object TNQTESEL: THLabel
      Left = 13
      Top = 9
      Width = 74
      Height = 13
      Caption = 'Total document'
    end
    object bRepartition: TToolbarButton97
      Left = 384
      Top = 21
      Width = 28
      Height = 27
      Hint = 'R'#233'partir la diff'#233'rence'
      DisplayMode = dmGlyphOnly
      Caption = 'R'#233'partition'
      Flat = False
      NumGlyphs = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = bRepartitionClick
      GlobalIndexImage = 'Z0041_S16G2'
    end
    object bsolde: TToolbarButton97
      Left = 354
      Top = 21
      Width = 28
      Height = 27
      Hint = 'Solder sur la ligne'
      DisplayMode = dmGlyphOnly
      Caption = 'Solde'
      Flat = False
      NumGlyphs = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = bsoldeClick
      GlobalIndexImage = 'Z0051_S16G2'
    end
    object NTotalDoc: THNumEdit
      Left = 95
      Top = 4
      Width = 85
      Height = 21
      TabStop = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      ReadOnly = True
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object NReste: THNumEdit
      Left = 95
      Top = 27
      Width = 85
      Height = 21
      TabStop = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 286
    Width = 418
    Height = 31
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 27
      ClientWidth = 418
      Caption = 'ToolWindow971'
      CloseButton = False
      ClientAreaHeight = 27
      ClientAreaWidth = 418
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      Resizable = False
      TabOrder = 0
      object bAbandon: TToolbarButton97
        Left = 352
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Alignment = taRightJustify
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object bValider: TToolbarButton97
        Left = 322
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Valider'
        Alignment = taRightJustify
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ModalResult = 1
        ParentShowHint = False
        ShowHint = True
        OnClick = bValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
      end
      object bAide: TToolbarButton97
        Left = 382
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        Alignment = taRightJustify
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object bsupprimer: TToolbarButton97
        Left = 3
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Supprimer la ligne en cours'
        DisplayMode = dmGlyphOnly
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bsupprimerClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
    end
  end
  object PCumul: THPanel
    Left = 0
    Top = 214
    Width = 418
    Height = 20
    Align = alBottom
    BevelOuter = bvNone
    Color = clWindow
    FullRepaint = False
    TabOrder = 4
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TCumul: TLabel
      Left = 8
      Top = 4
      Width = 30
      Height = 13
      Caption = 'Total'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object NTotDispo: THNumEdit
      Left = 52
      Top = -1
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object NTotSaisie: THNumEdit
      Left = 135
      Top = -1
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 240
    Top = 60
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;La somme des quantit'#233's saisies d'#233'passe la quantit'#233' t' +
        'otale ;W;O;O;O;'
      '1;?caption?;Vous devez renseigner une quantit'#233' positive;W;O;O;O;'
      '2;?caption?;Vous devez renseigner une quantit'#233' correcte;W;O;O;O;'
      
        '3;?caption?;La quantit'#233' totale n'#39'est pas r'#233'partie int'#233'gralement#' +
        '13 Voulez-vous continuer?;Q;YN;Y;N;'
      
        '4;?caption?;La quantit'#233' saisie est sup'#233'rieure '#224' la quantit'#233' disp' +
        'onible;W;O;O;O;'
      
        '5;?caption?;Les quantit'#233's disponibles sont insuffisantes;W;O;O;O' +
        ';'
      '6;?caption?;La quantit'#233' disponible est insuffisante;W;O;O;O;'
      '7;?caption?;Vous devez renseigner une date correcte;W;O;O;O;'
      '8;?caption?;Vous devez renseigner un num'#233'ro de lot;W;O;O;O;'
      '9;?caption?;Ce num'#233'ro lot existe d'#233'j'#224';W;O;O;O;'
      
        '10;?caption?;Certains lots n'#39'ont pas assez de quantit'#233's disponib' +
        'les.#13 Voulez-vous continuer?;Q;YN;Y;N;'
      
        '11;?caption?;Vous devez affecter les num'#233'ros de s'#233'rie de ce lot;' +
        'W;O;O;O;')
    Left = 292
    Top = 108
  end
  object MBTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Sortie de lots'
      'Qt'#233' sortie'
      'R'#233'ception de lots'
      'Qt'#233' r'#233'ceptionn'#233'e')
    Left = 228
    Top = 116
  end
end

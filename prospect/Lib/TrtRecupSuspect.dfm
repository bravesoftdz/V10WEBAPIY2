object FTrtRecupSuspect: TFTrtRecupSuspect
  Left = 235
  Top = 156
  Width = 635
  Height = 450
  HelpContext = 111000336
  Caption = 'Importation des Suspects'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pgeneral: TPanel
    Left = 0
    Top = 0
    Width = 619
    Height = 213
    Align = alTop
    Alignment = taLeftJustify
    TabOrder = 0
    DesignSize = (
      619
      213)
    object TRSS_PARSUSPECT: THLabel
      Left = 13
      Top = 20
      Width = 25
      Height = 13
      Caption = 'Code'
      FocusControl = RSS_PARSUSPECT
    end
    object TRSS_FICHIER: THLabel
      Left = 295
      Top = 20
      Width = 31
      Height = 13
      Caption = 'Fichier'
      FocusControl = RSS_FICHIER
    end
    object TRSS_SEPARATEUR: THLabel
      Left = 175
      Top = 85
      Width = 93
      Height = 13
      Caption = 'S'#233'parateur Champs'
      FocusControl = RSS_SEPARATEUR
    end
    object TRSS_FIN: THLabel
      Left = 496
      Top = 55
      Width = 63
      Height = 13
      Caption = 'Long. Champ'
    end
    object TRSS_DEBUT: THLabel
      Left = 384
      Top = 55
      Width = 67
      Height = 13
      Caption = 'Position d'#233'but'
    end
    object TRSS_LIBELLE: THLabel
      Left = 114
      Top = 20
      Width = 53
      Height = 13
      Caption = 'Description'
    end
    object TRSS_SEPTEXTE: THLabel
      Left = 384
      Top = 85
      Width = 59
      Height = 13
      Caption = 'Autre S'#233'par.'
      FocusControl = RSS_SEPTEXTE
    end
    object TRSS_PARTICULIER: THLabel
      Left = 175
      Top = 55
      Width = 91
      Height = 13
      Caption = 'Entrep. / Particulier'
    end
    object TRSS_LONGENREG: THLabel
      Left = 496
      Top = 85
      Width = 52
      Height = 13
      Hint = 'Longueur Enregistrement si Fixe'
      Caption = 'Long. ligne'
    end
    object RSS_PARSUSPECT: THCritMaskEdit
      Left = 48
      Top = 16
      Width = 55
      Height = 21
      Hint = 'Code du param'#233'trage'
      MaxLength = 10
      TabOrder = 0
      OnExit = RSS_PARSUSPECTExit
      TagDispatch = 0
    end
    object RSS_FICHIER: THCritMaskEdit
      Left = 344
      Top = 16
      Width = 246
      Height = 21
      Hint = 'Nom du Fichier R'#233'cup'#233'r'#233
      TabOrder = 2
      OnChange = RSS_FICHIERChange
      OnExit = RSS_FICHIERExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = RSS_FICHIERElipsisClick
    end
    object RSS_SEPARATEUR: THValComboBox
      Left = 274
      Top = 81
      Width = 102
      Height = 21
      ItemHeight = 13
      TabOrder = 7
      OnChange = RSS_SEPARATEURExit
      Items.Strings = (
        '')
      TagDispatch = 0
      Values.Strings = (
        '')
    end
    object RSS_Apercu: TMemo
      Left = 10
      Top = 111
      Width = 581
      Height = 95
      Anchors = [akLeft, akTop, akRight]
      Color = clMenu
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Courier New'
      Font.Pitch = fpFixed
      Font.Style = []
      Lines.Strings = (
        '')
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssBoth
      TabOrder = 10
      WordWrap = False
      OnClick = RSS_ApercuClick
      OnDblClick = BOffsetClick
    end
    object RSS_FormatFichier: TRadioGroup
      Left = 11
      Top = 44
      Width = 157
      Height = 61
      Caption = 'Format du fichier'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMenuText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        'Champ de longueur fixe'
        'D'#233'limit'#233' par un s'#233'parateur')
      ParentFont = False
      TabOrder = 3
      OnClick = RSS_FormatFichierClick
    end
    object RSS_DEBUT: THNumEdit
      Left = 459
      Top = 51
      Width = 31
      Height = 21
      TabStop = False
      Color = clMenu
      ReadOnly = True
      TabOrder = 5
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0'
      Masks.ZeroMask = '#,##0'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object RSS_FIN: THNumEdit
      Left = 566
      Top = 51
      Width = 31
      Height = 21
      TabStop = False
      Color = clInactiveBorder
      ReadOnly = True
      TabOrder = 6
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0'
      Masks.ZeroMask = '#,##0'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object RSS_LIBELLE: TMaskEdit
      Left = 175
      Top = 16
      Width = 113
      Height = 21
      MaxLength = 35
      TabOrder = 1
    end
    object RSS_SEPTEXTE: THCritMaskEdit
      Left = 459
      Top = 81
      Width = 31
      Height = 21
      MaxLength = 5
      TabOrder = 8
      TagDispatch = 0
    end
    object RSS_PARTICULIER: THValComboBox
      Left = 274
      Top = 51
      Width = 102
      Height = 21
      ItemHeight = 13
      TabOrder = 4
      OnExit = RSS_PARTICULIERExit
      Items.Strings = (
        'Entreprise'
        'Particulier')
      TagDispatch = 0
      Values.Strings = (
        'ENT'
        'PAR')
    end
    object RSS_LONGENREG: THNumEdit
      Left = 566
      Top = 81
      Width = 31
      Height = 21
      AutoSize = False
      Color = clInactiveBorder
      Enabled = False
      TabOrder = 9
      OnClick = RSS_ApercuClick
      Decimals = 0
      Digits = 6
      Masks.PositiveMask = '#,##0'
      Masks.ZeroMask = '#,##0'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object GCAPERCU: THGrid
      Left = 9
      Top = 108
      Width = 602
      Height = 99
      Anchors = [akLeft, akTop, akRight]
      ColCount = 1
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 1
      FixedRows = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 11
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      OnCellEnter = GcApercuCellEnter
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
    end
  end
  object PBOUTON: TDock97
    Left = 0
    Top = 374
    Width = 619
    Height = 38
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 610
      Caption = 'ToolWindow971'
      ClientAreaHeight = 34
      ClientAreaWidth = 610
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 386
        Top = 4
        Width = 29
        Height = 27
        Hint = 'Enregistrer le param'#233'trage'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
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
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 424
        Top = 4
        Width = 29
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 538
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
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
      object BNouveau: TToolbarButton97
        Left = 462
        Top = 4
        Width = 29
        Height = 27
        Hint = 'Nouveau param'#233'trage'
        Caption = 'Nouveau'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BNouveauClick
        GlobalIndexImage = 'Z2293_S24G1'
        IsControl = True
      end
      object bSupprimer: TToolbarButton97
        Left = 500
        Top = 4
        Width = 29
        Height = 27
        Hint = 'Suppression du param'#233'trage'
        Caption = 'Supprimer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BSupprimerClick
        GlobalIndexImage = 'Z0005_S16G1'
        IsControl = True
      end
    end
  end
  object PCorps: THPanel
    Left = 0
    Top = 213
    Width = 619
    Height = 161
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object PList: TPanel
      Left = 1
      Top = 1
      Width = 268
      Height = 159
      Align = alLeft
      AutoSize = True
      TabOrder = 0
      object TLChamp: THGrid
        Left = 1
        Top = 1
        Width = 266
        Height = 157
        Align = alLeft
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
        ScrollBars = ssVertical
        TabOrder = 0
        OnDblClick = BFlecheDroiteClick
        SortedCol = -1
        Titres.Strings = (
          'Nom des champs'
          '     Champs ')
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          124
          112)
      end
    end
    object PFleche: TPanel
      Left = 269
      Top = 1
      Width = 61
      Height = 159
      Align = alLeft
      TabOrder = 1
      object BFlecheDroite: TToolbarButton97
        Left = 18
        Top = 78
        Width = 28
        Height = 27
        Hint = 'Ajouter le champ'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BFlecheDroiteClick
        GlobalIndexImage = 'Z0056_S16G2'
        IsControl = True
      end
      object BFlecheGauche: TToolbarButton97
        Left = 18
        Top = 113
        Width = 28
        Height = 27
        Hint = 'Enlever le champ'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BFlecheGaucheClick
        GlobalIndexImage = 'Z0077_S16G2'
        IsControl = True
      end
      object Boffset: TToolbarButton97
        Left = 19
        Top = 24
        Width = 28
        Height = 27
        Hint = 'Affecter les valeurs position et longueur de l'#39'apercu fichier '
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BOffsetClick
        GlobalIndexImage = 'Z0609_S16G2'
        IsControl = True
      end
    end
    object PGRID: TPanel
      Left = 330
      Top = 1
      Width = 288
      Height = 159
      Align = alClient
      AutoSize = True
      TabOrder = 2
      object GChamp: THGrid
        Tag = 1
        Left = 1
        Top = 1
        Width = 286
        Height = 157
        Align = alClient
        ColCount = 8
        DefaultColWidth = 50
        DefaultRowHeight = 18
        RowCount = 35
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goTabs]
        TabOrder = 0
        SortedCol = 2
        Titres.Strings = (
          ''
          'Champs'
          'Position'
          'Longueur'
          'Code'
          'Formule'
          'Test'
          'Erreur')
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        OnCellEnter = GChampCellEnter
        OnCellExit = GChampCellExit
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        DBIndicator = True
        ColWidths = (
          10
          50
          50
          50
          50
          50
          50
          50)
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
          18)
      end
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 640
    Top = 300
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Le code du param'#232'trage est obligatoire;W;O;O;O;'
      '1;?caption?;Le code fournisseur est obligatoire;W;O;O;O;'
      '2;?caption?;Le code fournisseur n'#39'existe pas;W;O;O;O;'
      '3;?caption?;Le nom de fichier est obligatoire;W;O;O;O;'
      
        '4;?caption?;La longueur de l'#39'enregistrement est obligatoire;W;O;' +
        'O;O;'
      '5;?caption?;Le support est obligatoire;W;O;O;O;'
      '6;?caption?;Le s'#233'parateur est obligatoire;W;O;O;O;'
      '7;?caption?;Le texte de s'#233'paration est obligatoire;W;O;O;O;'
      
        '8;?caption?;Le type d'#39'article par d'#233'faut est obligatoire;W;O;O;O' +
        ';'
      
        '9;?caption?;La famille de taxe par d'#233'faut est obligatoire;W;O;O;' +
        'O;'
      '10;?caption?;Le champ n'#39'est pas s'#233'lectionnable;W;O;O;O;'
      '11;?caption?;Le champ s'#233'lectionn'#233' n'#39'existe pas;W;O;O;O;'
      
        '12;?caption?;La r'#233'f'#233'rence fournisseur doit '#234'tre dans la liste de' +
        's champs;W;O;O;O;'
      
        '13;?caption?;La famille comptable par d'#233'faut est obligatoire;W;O' +
        ';O;O;'
      '14;?caption?;La colonne doit etre diff'#233'rente de 0;W;O;O;O;'
      
        '15;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y' +
        ';C;'
      
        '16;?caption?;Etes-vous s'#251'r de vouloir supprimer ces donn'#233'es ?;Q;' +
        'YNC;Y;C;'
      
        '17;?caption?;La longueur totale des champs exc'#232'de celle de l'#39'enr' +
        'egistrement;W;O;O;O;'
      '18;?caption?;Fichier inaccessible;W;O;O;O;'
      '')
    Left = 676
    Top = 292
  end
  object OpenDialogButton: TOpenDialog
    Left = 556
    Top = 65532
  end
end

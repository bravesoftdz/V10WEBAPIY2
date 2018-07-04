object FSaisieAux: TFSaisieAux
  Left = 304
  Top = 232
  Width = 608
  Height = 372
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie balance Auxiliaire'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 592
    Height = 9
  end
  object DockRight: TDock97
    Left = 583
    Top = 9
    Width = 9
    Height = 293
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 293
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 302
    Width = 592
    Height = 31
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 502
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 502
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Valider la saisie'
        Caption = 'Valider'
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
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
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
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
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
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object Outils: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Calcul du solde'
        Caption = 'Solde'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeClick
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 33
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Rechercher dans la balance'
        Caption = 'Chercher'
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
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object ToolbarSep971: TToolbarSep97
        Left = 27
        Top = 0
      end
      object Sep97: TToolbarSep97
        Left = 87
        Top = 0
      end
      object BNew: TToolbarButton97
        Left = 60
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Nouveau Compte'
        Caption = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNewClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 574
    Height = 293
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object GS: THGrid
      Left = 0
      Top = 0
      Width = 574
      Height = 215
      Align = alClient
      BorderStyle = bsNone
      ColCount = 10
      Ctl3D = True
      DefaultRowHeight = 18
      Enabled = False
      FixedCols = 6
      RowCount = 500
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
      ParentCtl3D = False
      TabOrder = 0
      OnEnter = GSEnter
      OnExit = GSExit
      OnKeyPress = GSKeyPress
      OnMouseDown = GSMouseDown
      SortedCol = -1
      Titres.Strings = (
        'Etabl.'
        'Cr'#233#233'e'
        'Collectif'
        'TypeGene'
        'Lettrable'
        'ModeRegl'
        'Auxiliaire'
        'Libell'#233
        'D'#233'bit'
        'Cr'#233'dit')
      Couleur = False
      MultiSelect = False
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
      ColWidths = (
        2
        2
        2
        2
        2
        3
        92
        184
        144
        126)
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
    object PPied: THPanel
      Left = 0
      Top = 215
      Width = 574
      Height = 78
      Align = alBottom
      BevelOuter = bvNone
      BorderStyle = bsSingle
      Enabled = False
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object H_SOLDE: THLabel
        Left = 260
        Top = 38
        Width = 93
        Height = 13
        Caption = 'Solde Progressif'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel2: TBevel
        Left = 356
        Top = 8
        Width = 116
        Height = 17
      end
      object Bevel3: TBevel
        Left = 478
        Top = 8
        Width = 116
        Height = 17
      end
      object Bevel4: TBevel
        Left = 478
        Top = 37
        Width = 116
        Height = 17
      end
      object HConf: TToolbarButton97
        Left = 8
        Top = 50
        Width = 25
        Height = 24
        Hint = 'Ecriture confidentielle'
        Enabled = False
        NoBorder = True
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0141_S16G2'
      end
      object ISigneEuro: TImage
        Left = 34
        Top = 56
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
      object LSA_SOLDE: THLabel
        Left = 381
        Top = 37
        Width = 72
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TOTALCREDIT: THLabel
        Left = 137
        Top = 12
        Width = 117
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TOTALCREDIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TOTALDEBIT: THLabel
        Left = 8
        Top = 11
        Width = 108
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TOTALDEBIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object HLabel2: THLabel
        Left = 260
        Top = 10
        Width = 92
        Height = 13
        Caption = 'Totaux Comptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object FLASHDEVISE: TFlashingLabel
        Left = 63
        Top = 58
        Width = 46
        Height = 13
        Caption = 'DEVISE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object LSA_BALANCE: THLabel
        Left = 9
        Top = 29
        Width = 87
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_BALANCE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object HLabel1: THLabel
        Left = 32
        Top = 38
        Width = 83
        Height = 13
        Caption = 'Solde Balance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel7: TBevel
        Left = 125
        Top = 38
        Width = 116
        Height = 17
      end
      object SA_TOTALCREDIT: THNumEdit
        Tag = 1
        Left = 479
        Top = 8
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SOLDE: THNumEdit
        Tag = 1
        Left = 479
        Top = 40
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_TOTALDEBIT: THNumEdit
        Tag = 1
        Left = 357
        Top = 8
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_BALANCE: THNumEdit
        Tag = 1
        Left = 126
        Top = 40
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
  end
  object HBalance: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Balance non sold'#233'e;E;O;O;O;'
      '1;?caption?;Impossible de sauvegarder la balance;W;O;O;O;'
      
        '2;?caption?;Une saisie balance est en cours sur un autre poste;E' +
        ';O;O;O;'
      'Cumul Classe'
      'Chargement du plan comptable en cours'
      'Balance Au'
      'Perte'
      'B'#233'n'#233'fice'
      'Cr'#233'ation de compte interdit'
      'Cr'#233'ation de compte tiers interdit'
      'Exercices'
      
        '11;?caption?;Vous ne pouvez pas saisir de balance N-1 car il exi' +
        'ste une balance N-1 pour comparatif;E;O;O;O;'
      
        '12;?caption?;Aucun exercice de r'#233'f'#233'rence pour saisir la balance;' +
        'E;O;O;O;'
      
        '13;?caption?;Les A Nouveaux ont '#233't'#233' modifi'#233's (Lettrage et/ou Poi' +
        'ntage);E;O;O;O;'
      
        '14;?caption?;Vous ne pouvez pas saisir de balance N-1 car il exi' +
        'ste une balance N;E;O;O;O;'
      
        '15;?caption?;Vous ne pouvez pas saisir de balance pour comparati' +
        'f car il existe une balance N-1;E;O;O;O;'
      
        '16;?caption?;Des '#233'critures ont d'#233'j'#224' '#233't'#233' saisies sur l'#39'exercice;E' +
        ';O;O;O;'
      
        '17;?caption?;La date est en dehors de l'#39'exercice courant;E;O;O;O' +
        ';'
      '18;?caption?;Enregistrement des auxiliaires ?;Q;YN;Y;C;'
      'Choisir une balance de situation'
      'Aucune balance de situation'
      
        '21;?caption?;Confirmez-vous la supression de la balance ?;Q;YN;Y' +
        ';C;'
      
        '22;?caption?;Aucune '#233'criture pour la p'#233'riode s'#233'lectionn'#233'e;E;O;O;' +
        'O;'
      '23;?caption?;Veuillez param'#233'trer les comptes d'#39'attente;E;O;O;O;')
    Left = 103
    Top = 161
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 200
    Top = 161
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'BCreateExo'
      'BValEntete'
      'BDelBalance')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 152
    Top = 164
  end
end

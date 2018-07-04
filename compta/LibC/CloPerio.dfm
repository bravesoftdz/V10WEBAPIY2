object CloPer: TCloPer
  Left = 422
  Top = 179
  HelpContext = 7736000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Cl'#244'ture p'#233'riodique'
  ClientHeight = 254
  ClientWidth = 411
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 411
    Height = 212
    Align = alClient
    TabOrder = 0
    object HTexte1: THLabel
      Left = 12
      Top = 12
      Width = 389
      Height = 21
      AutoSize = False
      Caption = 'Ce traitement va contr'#244'ler l'#39'existence de mouvements'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object HLabel3: THLabel
      Left = 15
      Top = 41
      Width = 89
      Height = 13
      AutoSize = False
      Caption = '&P'#233'riode'
      FocusControl = FDateCpta1
    end
    object HLabel4: THLabel
      Left = 16
      Top = 16
      Width = 97
      Height = 13
      AutoSize = False
      Caption = 'Exercice'
      FocusControl = FExercice
    end
    object Htexte2: THLabel
      Left = 12
      Top = 36
      Width = 389
      Height = 21
      AutoSize = False
      Caption = 'autres que les mouvements "normaux" (simulation, etc...)'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object HDernClo: THLabel
      Left = 16
      Top = 72
      Width = 389
      Height = 20
      AutoSize = False
      Caption = 'Derni'#232're cloture p'#233'riodique :'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object HPatience: THLabel
      Left = 8
      Top = 117
      Width = 393
      Height = 16
      Alignment = taCenter
      AutoSize = False
      Caption = 'Traitement en cours...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object Resume: TMemo
      Left = 8
      Top = 140
      Width = 393
      Height = 65
      BorderStyle = bsNone
      Color = clMenu
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Lines.Strings = (
        
          'Attention : Il existe des '#233'critures de simulation ant'#233'rieures '#224' ' +
          'la date de cl'#244'ture.'
        
          'Attention : Il existe des '#233'critures de simulation ant'#233'rieures '#224' ' +
          'la date de cl'#244'ture.'
        
          'Attention : Il existe des '#233'critures de simulation ant'#233'rieures '#224' ' +
          'la date de cl'#244'ture.'
        
          'Attention : Il existe des '#233'critures de simulation ant'#233'rieures '#224' ' +
          'la date de cl'#244'ture.'
        
          'Attention : Il existe des '#233'critures de simulation ant'#233'rieures '#224' ' +
          'la date de cl'#244'ture.')
      ParentFont = False
      ReadOnly = True
      TabOrder = 3
    end
    object FExercice: THValComboBox
      Left = 121
      Top = 11
      Width = 241
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 0
      OnChange = FExerciceChange
      TagDispatch = 0
    end
    object FDateCpta1: THValComboBox
      Left = 125
      Top = 40
      Width = 128
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 1
      OnChange = FDateCpta1Change
      TagDispatch = 0
      Vide = True
    end
    object FDateCpta2: THValComboBox
      Left = 351
      Top = 56
      Width = 46
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 2
      Visible = False
      TagDispatch = 0
      Vide = True
    end
    object BArchive: TCheckBox
      Left = 16
      Top = 96
      Width = 241
      Height = 17
      Caption = 'Archivage automatique des donn'#233'es'
      Checked = True
      State = cbChecked
      TabOrder = 4
      OnClick = BArchiveClick
    end
    object FDateCptaExo: THValComboBox
      Left = 319
      Top = 64
      Width = 90
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 5
      Visible = False
      TagDispatch = 0
      Vide = True
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 212
    Width = 411
    Height = 42
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 38
      ClientWidth = 411
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 38
      ClientAreaWidth = 411
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 376
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 312
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Lancer la v'#233'rification'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
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
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 343
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
    end
  end
  object MsgBleme: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        'Attention : Il existe des '#233'critures de simulation ant'#233'rieures '#224' ' +
        'la date de cl'#244'ture.'
      
        'Attention : Il existe des '#233'critures de situation ant'#233'rieures '#224' l' +
        'a date de cl'#244'ture.'
      
        'Attention : Il existe des '#233'critures de pr'#233'vision ant'#233'rieures '#224' l' +
        'a date de cl'#244'ture.'
      
        'Attention : Il existe des '#233'critures de r'#233'visions ant'#233'rieures '#224' l' +
        'a date de cl'#244'ture.'
      
        'Attention : Il existe des '#233'critures d'#39'abonnement ant'#233'rieures '#224' l' +
        'a date de cl'#244'ture.'
      'Contr'#244'le avant cl'#244'ture '
      'Lancer la cl'#244'ture'
      'Contr'#244'le en cours. Veuillez patienter...'
      'Traitement en cours. Veuillez patienter...'
      'Derni'#232're cl'#244'ture p'#233'riodique :'
      'Lancer la d'#233'cl'#244'ture'
      '11')
    Left = 333
    Top = 156
  end
  object Confirmation: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;Cl'#244'ture p'#233'riodique;Confirmez-vous la cl'#244'ture ?;E;YN;Y;N'
      
        '1;Cl'#244'ture p'#233'riodique; La p'#233'riode que vous avez choisie est ant'#233'r' +
        'ieure '#224' la date de cl'#244'ture.;W;O;O;O'
      '2;;;E;0;0;0;'
      
        '3;Cl'#244'ture p'#233'riodique;Le traitement s'#39'est correctement termin'#233'.;E' +
        ';O;O;O;'
      
        '4;Contr'#244'le avant cl'#244'ture;Aucun mouvement particulier n'#39'a '#233't'#233' d'#233't' +
        'ect'#233'.;E;O;O;O;'
      
        '5;Contr'#244'le avant cl'#244'ture;Attention : des mouvements particuliers' +
        ' ont '#233't'#233' d'#233'tect'#233's.;E;O;O;O;'
      
        '6;Cl'#244'ture p'#233'riodique;ATTENTION.Confirmez-vous le traitement de c' +
        'l'#244'ture p'#233'riodique ?;Q;YN;N;N;'
      
        '7;D'#233'cl'#244'ture p'#233'riodique;ATTENTION.Confirmez-vous le traitement de' +
        ' d'#233'cl'#244'ture p'#233'riodique ?;Q;YN;N;N;'
      
        '8;D'#233'cl'#244'ture p'#233'riodique;Le traitement s'#39'est correctement termin'#233'.' +
        ';E;O;O;O;'
      
        '9;Cl'#244'ture p'#233'riodique;Vous devez cl'#244'turer p'#233'riodiquement toutes l' +
        'es p'#233'riodes de l'#39'exercice pr'#233'c'#233'dent.;E;O;O;O;'
      
        '10;D'#233'cl'#244'ture p'#233'riodique;Vous devez d'#233'cl'#244'turer p'#233'riodiquement tou' +
        'tes les p'#233'riodes de l'#39'exercice suivant.;E;O;O;O;'
      
        '11;Cl'#244'ture p'#233'riodique;ATTENTION.Pr'#233'sence d'#8217#233'critures non valid'#233'e' +
        's. Voulez-vous continuer ?;Q;YN;N;N;'
      ' ')
    Left = 371
    Top = 80
  end
  object Morceaux: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '2;Rappel apr'#232's validation; '
      'ATTENTION : Cette op'#233'ration est irr'#233'versible !'
      'Au pr'#233'alable :'
      '          - Assurez-vous d'#39'avoir pass'#233' toutes vos '#233'critures.'
      '          - Effectuez une sauvegarde.'
      'Confirmez-vous ?;E;YN;Y;N; ')
    Left = 368
    Top = 152
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 312
    Top = 76
  end
end

object FVerifCpte: TFVerifCpte
  Left = 316
  Top = 228
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Contr'#244'le des comptes'
  ClientHeight = 153
  ClientWidth = 381
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
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 381
    Height = 119
    Align = alClient
    TabOrder = 0
    object TFVerification: THLabel
      Left = 5
      Top = 11
      Width = 104
      Height = 13
      AutoSize = False
      Caption = '&V'#233'rification'
      FocusControl = FVerification
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Panel2: TPanel
      Left = 4
      Top = 36
      Width = 373
      Height = 75
      BevelOuter = bvLowered
      Caption = 'Panel2'
      TabOrder = 1
      object TNBError1: TLabel
        Left = 5
        Top = 2
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError2: TLabel
        Left = 5
        Top = 16
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError3: TLabel
        Left = 5
        Top = 30
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError5: TLabel
        Left = 5
        Top = 58
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError4: TLabel
        Left = 5
        Top = 44
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TCModR: THValComboBox
        Left = 332
        Top = 0
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 0
        Visible = False
        TagDispatch = 0
        DataType = 'TTMODEREGLE'
      end
      object TCRegTVA: THValComboBox
        Left = 332
        Top = 40
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 1
        Visible = False
        TagDispatch = 0
        DataType = 'TTREGIMETVA'
      end
      object TCTvaEnc: THValComboBox
        Left = 332
        Top = 60
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 2
        Visible = False
        TagDispatch = 0
        DataType = 'TTTVAENCAISSEMENT'
      end
      object TCCpte: TComboBox
        Left = 332
        Top = 24
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 3
        Text = 'TCCpte'
        Visible = False
      end
      object TCTypeConPart: THValComboBox
        Left = 284
        Top = 44
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 4
        Visible = False
        TagDispatch = 0
        DataType = 'TTTYPECONTREPARTIE'
      end
      object TCDev: THValComboBox
        Left = 240
        Top = 44
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 5
        Visible = False
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object TCnaGen: THValComboBox
        Left = 64
        Top = 8
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 6
        Visible = False
        TagDispatch = 0
        DataType = 'TTNATGENE'
      end
      object TCnaAux: THValComboBox
        Left = 64
        Top = 32
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 7
        Visible = False
        TagDispatch = 0
        DataType = 'TTNATTIERS'
      end
      object TCnaJal: THValComboBox
        Left = 64
        Top = 56
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 8
        Visible = False
        TagDispatch = 0
        DataType = 'TTNATJAL'
      end
      object TCRelTraite: THDBValComboBox
        Left = 122
        Top = 22
        Width = 41
        Height = 21
        Style = csDropDownList
        Color = clYellow
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 9
        Visible = False
        TagDispatch = 0
        DataType = 'TTRELANCETRAITE'
        DataField = 'T_RELANCETRAITE'
      end
      object TCRelRegle: THDBValComboBox
        Left = 198
        Top = 10
        Width = 41
        Height = 21
        Style = csDropDownList
        Color = clYellow
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 10
        Visible = False
        TagDispatch = 0
        DataType = 'TTRELANCEREGLE'
        DataField = 'T_RELANCEREGLEMENT'
      end
      object CbEcr: THValComboBox
        Left = 272
        Top = 16
        Width = 37
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 11
        Visible = False
        TagDispatch = 0
      end
    end
    object FVerification: TComboBox
      Left = 132
      Top = 8
      Width = 245
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      Items.Strings = (
        'Toutes'
        'G'#233'n'#233'raux'
        'Tiers'
        'Section'
        'Journaux')
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 119
    Width = 381
    Height = 34
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 1
    object TTravail: TLabel
      Left = 5
      Top = 10
      Width = 241
      Height = 13
      AutoSize = False
      Caption = 'TTravail'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object BValider: THBitBtn
      Left = 318
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Lancer la v'#233'rification'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValiderClick
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
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 349
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Glyph.Data = {
        CE070000424DCE07000000000000360000002800000024000000120000000100
        1800000000009807000000000000000000000000000000000000008284008284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284008284008284008284008284008284008284008284
        0082840082840082840082840082840082848482848482840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284FFFFFF008284008284008284008284008284008284
        0082840082840082840082840082840082840082840082840082840000FF0000
        840000848482840082840082840082840082840082840000FF84828400828400
        8284008284008284008284008284008284008284848284848284FFFFFF008284
        008284008284008284008284008284FFFFFF0082840082840082840082840082
        840082840082840000FF00008400008400008484828400828400828400828400
        00FF000084000084848284008284008284008284008284008284008284848284
        FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
        FF0082840082840082840082840082840082840000FF00008400008400008400
        00848482840082840000FF000084000084000084000084848284008284008284
        008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
        FF848284008284008284848284FFFFFF00828400828400828400828400828400
        82840000FF000084000084000084000084848284000084000084000084000084
        000084848284008284008284008284008284008284848284FFFFFF0082840082
        84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
        82840082840082840082840082840082840000FF000084000084000084000084
        0000840000840000840000848482840082840082840082840082840082840082
        84008284848284FFFFFF00828400828400828484828400828400828400828400
        8284FFFFFF848284008284008284008284008284008284008284008284008284
        0000FF0000840000840000840000840000840000848482840082840082840082
        84008284008284008284008284008284008284848284FFFFFF00828400828400
        8284008284008284008284FFFFFF848284008284008284008284008284008284
        0082840082840082840082840082840000840000840000840000840000848482
        8400828400828400828400828400828400828400828400828400828400828400
        8284848284FFFFFF008284008284008284008284008284848284008284008284
        0082840082840082840082840082840082840082840082840082840000FF0000
        8400008400008400008484828400828400828400828400828400828400828400
        8284008284008284008284008284008284848284FFFFFF008284008284008284
        8482840082840082840082840082840082840082840082840082840082840082
        840082840000FF00008400008400008400008400008484828400828400828400
        8284008284008284008284008284008284008284008284008284008284848284
        008284008284008284008284848284FFFFFF0082840082840082840082840082
        840082840082840082840082840000FF00008400008400008484828400008400
        0084000084848284008284008284008284008284008284008284008284008284
        008284008284848284008284008284008284008284008284848284FFFFFF0082
        840082840082840082840082840082840082840082840000FF00008400008400
        00848482840082840000FF000084000084000084848284008284008284008284
        008284008284008284008284008284848284008284008284008284848284FFFF
        FF008284008284848284FFFFFF00828400828400828400828400828400828400
        82840000FF0000840000848482840082840082840082840000FF000084000084
        000084848284008284008284008284008284008284008284848284FFFFFF0082
        84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
        82840082840082840082840082840082840000FF000084008284008284008284
        0082840082840000FF0000840000840000840082840082840082840082840082
        84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
        8284008284848284FFFFFF008284008284008284008284008284008284008284
        0082840082840082840082840082840082840082840000FF0000840000FF0082
        8400828400828400828400828400828400828484828484828400828400828400
        8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284008284008284008284008284848284848284848284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284008284008284008284008284008284008284008284
        008284008284008284008284008284008284}
      NumGlyphs = 2
    end
    object BStop: THBitBtn
      Left = 287
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Arr'#234'te la v'#233'rification'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BStopClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0107_S16G1'
      IsControl = True
    end
  end
  object MsgBar: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Analyse des comptes en cours'
      'Compte'
      'Libell'#233
      'Axe'
      'erreurs'
      'erreur'
      'Aucune erreur'
      'G'#233'n'#233'raux'
      'Tiers'
      'Sections'
      '10 ;'
      'Pr'#233'paration des fichiers...'
      'Code'
      'Journaux'
      'R'#233'paration des comptes'
      'V'#233'rification de comptes')
    Left = 88
    Top = 3
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Aucun enregistrement ne correspond aux crit'#232'res que ' +
        'vous avez s'#233'lectionn'#233's.;W;O;O;O;'
      '1;?caption?;Aucune erreur n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;'
      
        '2;?caption?;Confirmez-vous l'#39'arr'#234't de la v'#233'rification en cours ?' +
        ';Q;YN;N;N;'
      '3;?caption?;La r'#233'paration a '#233't'#233' effectu'#233'e;E;O;O;O;'
      
        '4;R'#233'paration des comptes; erreur(s) impossible(s) '#224' r'#233'parer car ' +
        'le compte est mouvement'#233';E;O;O;O;')
    Left = 56
    Top = 3
  end
  object QCG: TQuery
    DatabaseName = 'SOC'
    Left = 218
    Top = 4
  end
  object MsgLibel: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Le code n'#39'est pas renseign'#233
      'La longueur du code est non conforme'
      'L'#39'axe n'#39'est pas renseign'#233
      
        'La longueur du code est non conforme sur cet axe, le maximum est' +
        ' de'
      'Le libell'#233' n'#39'est pas renseign'#233' '
      'La nature du compte n'#39'est pas renseign'#233'e'
      'La nature du compte n'#39'existe pas'
      'Incoh'#233'rence entre la nature du compte et les caract'#233'ristiques'
      'Caract'#233'ristiques du compte incoh'#233'rentes'
      'Caract'#232're du compte non autoris'#233
      '10 ;'
      'Le compte de correspondance n'#176'1 n'#39'existe pas'
      'Le compte de correspondance n'#176'2 n'#39'existe pas'
      'La zone G_TVA ne devrait pas '#234'tre renseign'#233
      'La zone G_TVASURENCAISSEMENT ne devrait pas '#234'tre renseign'#233'e'
      'La zone G_TPF ne devrait pas '#234'tre renseign'#233'e'
      
        'L'#39'info pays du g'#233'n'#233'ral tiers (cr'#233'diteur ou d'#233'biteur) n'#39'existe pa' +
        's'
      
        'Le r'#233'gime TVA du compte g'#233'n'#233'ral de nature tiers cr'#233'diteur ou d'#233'b' +
        'iteur n'#39'existe pas'
      
        'La TVA sur encaissement du compte g'#233'n'#233'ral de nature tiers cr'#233'dit' +
        'eur n'#39'est pas renseign'#233'e'
      
        'La TVA sur encaissement du compte g'#233'n'#233'ral de nature tiers cr'#233'dit' +
        'eur ou d'#233'biteur n'#39'existe pas'
      
        'Le mode de r'#232'glement du compte g'#233'n'#233'ral de nature tiers cr'#233'diteur' +
        ' ou d'#233'biteur n'#39'est pas renseign'#233
      
        'Le mode de r'#232'glement du compte g'#233'n'#233'ral de nature tiers cr'#233'diteur' +
        ' ou d'#233'biteur n'#39'existe pas'
      '22 ;'
      'Le qualifiant quantit'#233' n'#176'1 n'#39'existe pas'
      'Le qualifiant quantit'#233'  n'#176'2 n'#39'existe pas'
      'Le compte g'#233'n'#233'ral collectif n'#39'est pas renseign'#233
      'Le compte g'#233'n'#233'ral collectif n'#39'existe pas'
      
        'Nature du compte auxiliaire incoh'#233'rente avec celle du compte g'#233'n' +
        #233'ral collectif'
      'Le mode de r'#232'glement n'#39'est pas renseign'#233
      'Le mode de r'#232'glement n'#39'existe pas'
      'L'#39' axe n'#39'existe pas'
      
        'Le type de relance de r'#232'glement du compte g'#233'n'#233'ral de nature tier' +
        's cr'#233'diteur ou d'#233'biteur  n'#39'existe pas'
      
        'Le type de relance de traite du compte g'#233'n'#233'ral de nature  tiers ' +
        'cr'#233'diteur ou d'#233'biteur n'#39'existe pas'
      '33 ;'
      
        'Le compte de nature banque n'#39'a pas de devise renseign'#233'e dans la ' +
        'fiche Compte Bancaire'
      
        'Le compte de nature banque a une devise qui n'#39'existe pas dans la' +
        ' fiche Compte Bancaire'
      'Compte non-lettrable avec un dernier code lettrage renseign'#233
      
        'Le r'#233'gime TVA du compte g'#233'n'#233'ral de nature tiers cr'#233'diteur ou d'#233'b' +
        'iteur lettrable n'#39'est pas renseign'#233
      'Le r'#233'gime TVA n'#39'est pas renseign'#233
      'Ce r'#233'gime TVA n'#39'existe pas'
      'La TVA encaissement n'#39'est pas renseign'#233'e'
      'La TVA encaissement n'#39'existe pas'
      'Le type de relance de r'#232'glement n'#39'existe pas'
      'Le type de relance de traite n'#39'existe pas'
      '44 ;'
      
        'Relev'#233' de facture impossible sur un tiers lettrable autre que cl' +
        'ient ou d'#233'biteur divers'
      'Le compteur saisi en facturier normal n'#39'existe pas'
      'Le compteur saisi en facturier simulation n'#39'existe pas'
      'Le compte de contrepartie n'#39'est pas renseign'#233
      'Le compte de contrepartie n'#39'existe pas'
      'Le compte de contrepartie n'#39'est pas de nature banque'
      'Le compte de contrepartie n'#39'est pas de nature caisse'
      'Le type de contrepartie n'#39'est pas renseign'#233
      'Le type de contrepartie n'#39'existe pas'
      'Le compte de contrepartie ne devrait pas '#234'tre renseign'#233
      'Le type de contrepartie ne devrait pas '#234'tre renseign'#233
      '56 ;'
      'Le compte de r'#233'gularisation ne devrait pas '#234'tre renseign'#233
      'Le compte de r'#233'gularisation n'#39'existe pas')
    Left = 123
    Top = 3
  end
  object MsgLibel2: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Erreur : Champ'
      'incorrect'
      'R'#233'gularisation de l'#39#233'criture')
    Left = 155
    Top = 3
  end
  object MsgCor: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Left = 187
    Top = 3
  end
  object QCT: TQuery
    DatabaseName = 'SOC'
    Left = 250
    Top = 4
  end
  object QCS: TQuery
    DatabaseName = 'SOC'
    Left = 282
  end
  object QCJ: TQuery
    DatabaseName = 'SOC'
    Left = 310
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 44
  end
end

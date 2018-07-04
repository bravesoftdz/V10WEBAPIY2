object FCalcScoring: TFCalcScoring
  Left = 265
  Top = 168
  HelpContext = 7583000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Calcul du scoring'
  ClientHeight = 170
  ClientWidth = 353
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
  object Blabla: TLabel
    Left = 4
    Top = 78
    Width = 346
    Height = 55
    AutoSize = False
    Caption = 
      'Le scoring permet de conna'#238'tre le niveau de relance moyen d'#39'un t' +
      'iers, calcul'#233' suivant le nombre de relances avant paiement des '#233 +
      'ch'#233'ances. Il ne prend en compte que les '#233'critures totalement let' +
      'tr'#233'es des tiers (clients ou autres d'#233'biteurs) lettrables.'
    ShowAccelChar = False
    WordWrap = True
  end
  object Dock: TDock97
    Left = 0
    Top = 134
    Width = 353
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 353
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 353
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BStop: TToolbarButton97
        Tag = 1
        Left = 223
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Arr'#234'ter le calcul du scoring'
        DisplayMode = dmGlyphOnly
        Caption = 'Stopper'
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
        OnClick = BStopClick
        GlobalIndexImage = 'Z0107_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 255
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Lancer le traitement'
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
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 287
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
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
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
      end
      object BAide: TToolbarButton97
        Left = 319
        Top = 3
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
    end
  end
  object PSelect: TGroupBox
    Left = 0
    Top = 0
    Width = 353
    Height = 77
    Align = alTop
    Caption = ' S'#233'lection '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object EMin: TLabel
      Left = 8
      Top = 48
      Width = 275
      Height = 13
      Caption = 'Nombre minimum d'#39#233'critures totalement lettr'#233'es pour calcul'
      FocusControl = Min
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label1: TLabel
      Left = 8
      Top = 20
      Width = 56
      Height = 13
      Caption = '&Auxiliaire de'
      FocusControl = AUXI
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 208
      Top = 20
      Width = 12
      Height = 13
      Caption = 'au'
      FocusControl = AUXI_
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object AUXI: THCpteEdit
      Left = 84
      Top = 16
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 17
      ParentFont = False
      TabOrder = 0
      ZoomTable = tzTToutDebit
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object AUXI_: THCpteEdit
      Left = 224
      Top = 16
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 17
      ParentFont = False
      TabOrder = 1
      ZoomTable = tzTToutDebit
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object Min: TSpinEdit
      Left = 300
      Top = 44
      Width = 45
      Height = 22
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxValue = 999
      MinValue = 3
      ParentFont = False
      TabOrder = 2
      Value = 3
    end
  end
  object Patience: TPanel
    Left = 0
    Top = 77
    Width = 353
    Height = 57
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    Visible = False
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 353
      Height = 57
      Align = alClient
    end
    object EnCours: TLabel
      Left = 4
      Top = 3
      Width = 182
      Height = 13
      Caption = 'Calcul en cours pour le tiers ...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsItalic]
      ParentFont = False
    end
    object LibAuxi: TLabel
      Left = 4
      Top = 19
      Width = 253
      Height = 13
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object NbTiers: TLabel
      Left = 4
      Top = 40
      Width = 141
      Height = 13
      Caption = 'Nombre de tiers trait'#233's : '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Anim: TAnimate
      Left = 264
      Top = 4
      Width = 48
      Height = 45
      CommonAVI = aviFindComputer
      StopFrame = 8
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Calcul du scoring de relance;Vous devez s'#233'lectionner une fourc' +
        'hette de comptes auxiliaires.;W;O;O;O;'
      
        '1;Calcul du scoring de relance;Aucun enregistrement ne correspon' +
        'd '#224' votre s'#233'lection.;E;O;O;O;'
      
        '2;Calcul du scoring de relance;D'#233'sirez-vous un compte-rendu du c' +
        'alcul du scoring ?;Q;YNC;N;C;'
      
        '3;Calcul du scoring de relance;D'#233'sirez-vous lancer le traitement' +
        ' ?;Q;YNC;N;C;'
      
        '4;Calcul du scoring de relance;Confirmez-vous l'#39'arr'#234't du traitem' +
        'ent en cours ?;Q;YN;N;N;'
      'Nombre de tiers trait'#233's : ')
    Left = 264
    Top = 4
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 216
    Top = 84
  end
end

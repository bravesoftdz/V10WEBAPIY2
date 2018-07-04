object FCtrlDiv: TFCtrlDiv
  Left = 411
  Top = 117
  HelpContext = 3215000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'R'#233'paration des fichiers'
  ClientHeight = 451
  ClientWidth = 284
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 284
    Height = 411
    Align = alClient
    TabOrder = 0
    object TTravail: TLabel
      Left = 10
      Top = 364
      Width = 181
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
    object TTravailLeq: TLabel
      Left = 10
      Top = 393
      Width = 181
      Height = 14
      AutoSize = False
      Caption = 'TTravailLeq'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object FControleMvt: TGroupBox
      Left = 7
      Top = 131
      Width = 270
      Height = 259
      Caption = '  Mouvements  '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object FDatePaquet: TCheckBox
        Left = 12
        Top = 32
        Width = 246
        Height = 19
        Alignment = taLeftJustify
        Caption = '&Dates de paquet lettr'#233
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 1
      end
      object FNumPL: TCheckBox
        Left = 12
        Top = 67
        Width = 246
        Height = 14
        Alignment = taLeftJustify
        Caption = '&Num'#233'ro de pi'#232'ce/ligne'
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 3
      end
      object FLettrage: TCheckBox
        Left = 12
        Top = 50
        Width = 246
        Height = 13
        Alignment = taLeftJustify
        Caption = '&Lettrage'
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 2
      end
      object FContreparties: TCheckBox
        Left = 12
        Top = 85
        Width = 246
        Height = 13
        Alignment = taLeftJustify
        Caption = '&Contreparties'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 4
      end
      object FReparMvt: TCheckBox
        Left = 12
        Top = 19
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Equilibrage des mouvements'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 0
      end
      object FDevMnt: TCheckBox
        Left = 12
        Top = 102
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Montants Frcs/Euro sur pi'#232'ces en &devise'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 5
      end
      object FPeriode: TCheckBox
        Left = 12
        Top = 120
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Recalcul des p'#233'riodes et semaines'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 6
      end
      object FLettRef: TCheckBox
        Left = 12
        Top = 137
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Lettrage sur &code de regroupement'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 7
      end
      object FTvaEnc: TCheckBox
        Left = 12
        Top = 155
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Recalcul des &bases de tva sur encaissement'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 8
      end
      object FCodeAccept: TCheckBox
        Left = 12
        Top = 172
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'R'#233'initialisation code acceptation'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 9
      end
      object FSoldeP: TCheckBox
        Left = 12
        Top = 190
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Recalcul totaux &point'#233's '
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 10
        Visible = False
      end
      object FEntetePiece: TCheckBox
        Left = 12
        Top = 208
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'R'#233'paration des &en-t'#234'tes de pi'#232'ces'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 11
      end
      object ANNREVISION: TCheckBox
        Left = 12
        Top = 224
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Annulation du dernier envoi de r'#233'vision'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 12
      end
      object ckAuxiAna: TCheckBox
        Left = 12
        Top = 240
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'R'#233'paration des auxiliaires sur l'#39'analytique'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 13
      end
    end
    object FcontroleCpt: TGroupBox
      Left = 7
      Top = 6
      Width = 270
      Height = 119
      Caption = '  Fiches  '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object FGen: TCheckBox
        Left = 11
        Top = 16
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Comptes &g'#233'n'#233'raux'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 0
      end
      object FAux: TCheckBox
        Left = 11
        Top = 32
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Comptes &auxiliaires'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 1
      end
      object FSec: TCheckBox
        Left = 11
        Top = 49
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Sections'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 2
      end
      object FJal: TCheckBox
        Left = 11
        Top = 65
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Journaux'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 3
      end
      object FSolde: TCheckBox
        Left = 11
        Top = 82
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Recalcul des soldes des comptes'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 4
      end
      object FANouveau: TCheckBox
        Left = 11
        Top = 98
        Width = 246
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Recalcul des '#224'-nouveaux'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 5
      end
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 411
    Width = 284
    Height = 40
    AllowDrag = False
    BoundLines = [blTop, blBottom, blLeft, blRight]
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 400
      Caption = 'Actions'
      ClientAreaHeight = 34
      ClientAreaWidth = 400
      DockPos = 0
      HideWhenInactive = False
      MinClientHeight = 30
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 252
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
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
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 195
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Lancer la v'#233'rification'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'V'#233'rifier'
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
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 223
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
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;R'#233'paration des fichiers;Votre s'#233'lection est incorrecte;W;O;O;O' +
        ';'
      
        '1;R'#233'paration des fichiers;La r'#233'paration des fichiers est termin'#233 +
        'e;A;O;O;O;'
      
        '2;R'#233'paration des fichiers;Confirmez-vous la r'#233'paration des fichi' +
        'ers ?;Q;YN;N;N;'
      
        '3;R'#233'paration des fichiers; erreur ne peut pas '#234'tre r'#233'par'#233'e car l' +
        'e compte est mouvement'#233';E;O;O;O;'
      
        '4;R'#233'paration des fichiers; erreurs ne peuvent pas '#234'tre r'#233'par'#233'es ' +
        '"compte(s) mouvement'#233'(s)";E;O;O;O;')
    Left = 116
    Top = 15
  end
  object Mes: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'G'#233'n'#233'raux'
      'Tiers'
      'Sections'
      'Journaux'
      'Analyse en cours...')
    Left = 152
    Top = 15
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 160
    Top = 364
  end
  object MsgBar: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'R'#233'paration en cours...'
      'des g'#233'n'#233'raux'
      'des tiers'
      'des sections'
      'des journaux'
      'des soldes des comptes'
      'des '#224'-nouveaux'
      'des dates de paquet lettr'#233
      'des mouvements lettr'#233's'
      'des num'#233'ros de pi'#232'ce/ligne'
      'de l'#39'euro'
      'des contreparties'
      'de l'#39#233'quilibre comptable'
      'des p'#233'riodes et semaines'
      'des montants euros des A-Nouveaux')
    Left = 208
    Top = 223
  end
end

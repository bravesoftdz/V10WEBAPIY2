object FCreRubBu: TFCreRubBu
  Left = 265
  Top = 159
  HelpContext = 15411000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Cr'#233'ation des rubriques budg'#233'taires'
  ClientHeight = 311
  ClientWidth = 278
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
  object Dock: TDock97
    Left = 0
    Top = 275
    Width = 278
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 278
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 278
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 180
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Valider'
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
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 212
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 244
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
    end
  end
  object PFen: TPanel
    Left = 0
    Top = 0
    Width = 278
    Height = 275
    Align = alClient
    Caption = ' '
    TabOrder = 1
    object Label1: TLabel
      Left = 8
      Top = 16
      Width = 34
      Height = 13
      Caption = '&Budget'
      FocusControl = CbJal
    end
    object GbChoix: TGroupBox
      Left = 4
      Top = 126
      Width = 271
      Height = 143
      Caption = ' Choix des cr'#233'ations '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      object Patience: TLabel
        Left = 74
        Top = 119
        Width = 117
        Height = 13
        AutoSize = False
        Caption = 'Veuillez patienter...'
        FocusControl = CbJal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object Cb1: TCheckBox
        Left = 10
        Top = 24
        Width = 250
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Rubriques simples des comptes budg'#233'taires'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object Cb2: TCheckBox
        Left = 10
        Top = 49
        Width = 250
        Height = 17
        Alignment = taLeftJustify
        Caption = 'R&ubriques simples des sections budg'#233'taires'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object Cb3: TCheckBox
        Left = 10
        Top = 74
        Width = 250
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Rubr&iques composites comptes / sections'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object Cb4: TCheckBox
        Left = 10
        Top = 99
        Width = 250
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Rubri&ques composites sections / comptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
      end
    end
    object CbJal: THValComboBox
      Left = 54
      Top = 12
      Width = 221
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = CbJalChange
      TagDispatch = 0
      DataType = 'TTBUDJAL'
    end
    object CbRubG: THValComboBox
      Left = 12
      Top = 132
      Width = 41
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 3
      Visible = False
      TagDispatch = 0
      Vide = True
    end
    object CbRubS: THValComboBox
      Left = 52
      Top = 128
      Width = 41
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 4
      Visible = False
      TagDispatch = 0
      Vide = True
    end
    object RgChoix: TRadioGroup
      Left = 4
      Top = 59
      Width = 271
      Height = 63
      Caption = ' &Type de cr'#233'ation '
      ItemIndex = 0
      Items.Strings = (
        'A partir des &mouvements budg'#233'taires'
        'A partir des &croisements comptes/sections')
      TabOrder = 1
    end
    object CbDelTout: TCheckBox
      Left = 8
      Top = 40
      Width = 263
      Height = 17
      Alignment = taLeftJustify
      Caption = '&Supprimer la totalit'#233' des rubriques du budget O/N'
      TabOrder = 5
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Cr'#233'ation des rubriques budg'#233'taires;Aucun journal n'#39'a '#233't'#233' s'#233'lec' +
        'tionn'#233';W;O;O;O;'
      
        '1;Cr'#233'ation des rubriques budg'#233'taires;Aucun type de cr'#233'ation de r' +
        'ubrique n'#39'a '#233't'#233' s'#233'lectionn'#233';W;O;O;O;'
      
        '2;Cr'#233'ation des rubriques budg'#233'taires;ATTENTION ! Toutes les rubr' +
        'iques budg'#233'taires existantes vont '#234'tre d'#233'truites et remplac'#233'es. ' +
        'Confirmez-vous le traitement?;Q;YN;N;N;'
      
        '3;Cr'#233'ation des rubriques budg'#233'taires;Des codes rubriques n'#39'ont p' +
        'as '#233't'#233' cr'#233#233's car ils existaient d'#233'j'#224'.;W;O;O;O;'
      
        '4;Cr'#233'ation des rubriques budg'#233'taires;Des codes rubriques n'#39'ont p' +
        'as '#233't'#233' cr'#233#233's parce que certains comptes n'#39'ont pas de code rubriq' +
        'ue renseign'#233'. Vous devez renseigner le code rubrique des comptes' +
        ' budg'#233'taires g'#233'n'#233'raux ou des sections.;W;O;O;O;'
      'Traitement en cours'
      
        '6;Cr'#233'ation des rubriques budg'#233'taires;ATTENTION ! Les rubriques b' +
        'udg'#233'taires concern'#233'es vont '#234'tre d'#233'truites et remplac'#233'es. Confirm' +
        'ez-vous le traitement?;Q;YN;N;N;')
    Left = 28
    Top = 264
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 191
    Top = 133
  end
end

object ParamClo: TParamClo
  Left = 471
  Top = 107
  HelpContext = 7751100
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  ClientHeight = 210
  ClientWidth = 466
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GParamClo: TPanel
    Left = 0
    Top = 0
    Width = 466
    Height = 210
    Align = alClient
    TabOrder = 0
    object GroupBox3: TGroupBox
      Left = 9
      Top = 4
      Width = 449
      Height = 81
      Caption = 'Compte de bilan cl'#244'ture...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object HLabel7: THLabel
        Left = 21
        Top = 24
        Width = 130
        Height = 13
        Caption = '&G'#233'n'#233'ration compte de bilan'
        FocusControl = CloContrep
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HCloPiece: THLabel
        Left = 21
        Top = 52
        Width = 82
        Height = 13
        Caption = '&Pi'#232'ces g'#233'n'#233'r'#233'es '
        FocusControl = CloPiece
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object CloContrep: THValComboBox
        Left = 194
        Top = 20
        Width = 246
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTCLOCONTREPARTIE'
      end
      object CloPiece: THValComboBox
        Left = 194
        Top = 48
        Width = 246
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        OnChange = CloPieceChange
        TagDispatch = 0
        DataType = 'TTCLOPIECE'
      end
    end
    object GroupBox4: TGroupBox
      Left = 9
      Top = 87
      Width = 449
      Height = 85
      Caption = 'Compte de bilan ouverture...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object HLabel11: THLabel
        Left = 19
        Top = 24
        Width = 130
        Height = 13
        Caption = '&G'#233'n'#233'ration compte de bilan'
        FocusControl = AnoContrep
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HAnoPiece: THLabel
        Left = 19
        Top = 52
        Width = 82
        Height = 13
        Caption = '&Pi'#232'ces g'#233'n'#233'r'#233'es '
        FocusControl = AnoPiece
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object AnoContrep: THValComboBox
        Left = 193
        Top = 20
        Width = 247
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTCLOCONTREPARTIE'
      end
      object AnoPiece: THValComboBox
        Left = 193
        Top = 48
        Width = 247
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        OnChange = AnoPieceChange
        TagDispatch = 0
        DataType = 'TTCLOPIECE'
      end
    end
    object GBPVBQE: TGroupBox
      Left = 9
      Top = 174
      Width = 449
      Height = 155
      Caption = 'Comptes pointables... '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      Visible = False
      object ANOCPTBQ: TRadioGroup
        Left = 7
        Top = 18
        Width = 435
        Height = 52
        Caption = 'Comptes pointables de nature banque... '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemIndex = 1
        Items.Strings = (
          'Un mouvement pour le &total point'#233' de l'#39'exercice'
          'Un mouvement par &r'#233'f'#233'rence de pointage de l'#39'exercice')
        ParentFont = False
        TabOrder = 0
      end
      object ANOCptPV: TRadioGroup
        Left = 7
        Top = 77
        Width = 435
        Height = 68
        Caption = 'Comptes pointables et ventilables autres que banque...  '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemIndex = 0
        Items.Strings = (
          'Reprise en solde avec &ventilation analytique'
          
            'Un mouvement par r'#233'f'#233'rence de &pointage avec ventilation sur sec' +
            'tion d'#39'attente'
          
            'Un mouvement pour le t&otal point'#233' de l'#39'exercice avec ventilatio' +
            'n sur section d'#39'attente')
        ParentFont = False
        TabOrder = 1
      end
    end
    object Panel1: TPanel
      Left = 1
      Top = 174
      Width = 464
      Height = 35
      Align = alBottom
      BevelInner = bvLowered
      TabOrder = 3
      object BCValide: THBitBtn
        Tag = 1
        Left = 368
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Lancer la cl'#244'ture'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BCValideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BCAide: THBitBtn
        Tag = 1
        Left = 432
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BCAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 400
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
    end
  end
  object HMess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Param'#233'trage de la cl'#244'ture provisoire'
      'Param'#233'trage de la cl'#244'ture d'#233'finitive')
    Left = 332
    Top = 65532
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 386
    Top = 140
  end
end

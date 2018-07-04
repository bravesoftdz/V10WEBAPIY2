object FTvaLettr: TFTvaLettr
  Left = 253
  Top = 155
  HelpContext = 7648100
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 
    'Rapprochement des bases d'#39'encaissement de l'#39'acompte et des factu' +
    'res'
  ClientHeight = 314
  ClientWidth = 534
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
  object POutils: TPanel
    Left = 0
    Top = 281
    Width = 534
    Height = 33
    Align = alBottom
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 0
    object BValide: THBitBtn
      Left = 436
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Valider'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0184_S16G1'
      IsControl = True
    end
    object BAbandon: THBitBtn
      Left = 468
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ModalResult = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BAide: THBitBtn
      Left = 500
      Top = 3
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
      TabOrder = 2
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BZoomFact: THBitBtn
      Tag = 1
      Left = 4
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Voir facture'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BZoomFactClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
    object BZoomLettre: THBitBtn
      Tag = 1
      Left = 36
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Voir lettrage'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = BZoomLettreClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0212_S16G1'
      IsControl = True
    end
    object BImprimer: THBitBtn
      Left = 404
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Caption = 'BImprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
  end
  object PFen: TPanel
    Left = 0
    Top = 0
    Width = 534
    Height = 281
    Align = alClient
    TabOrder = 1
    object Bevel2: TBevel
      Left = 1
      Top = 103
      Width = 532
      Height = 177
      Align = alTop
    end
    object Bevel3: TBevel
      Left = 1
      Top = 1
      Width = 532
      Height = 102
      Align = alTop
    end
    object Label4: TLabel
      Left = 6
      Top = 58
      Width = 83
      Height = 13
      Caption = 'Montant acompte'
    end
    object Label6: TLabel
      Left = 6
      Top = 26
      Width = 89
      Height = 13
      Caption = 'Total TTC factures'
    end
    object Label8: TLabel
      Left = 53
      Top = 108
      Width = 96
      Height = 13
      Caption = 'Couverture acompte'
    end
    object Label9: TLabel
      Left = 196
      Top = 108
      Width = 124
      Height = 13
      Caption = 'Total couvertures factures'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 412
      Top = 108
      Width = 62
      Height = 13
      Caption = 'Ratio calcul'#233
    end
    object Label10: TLabel
      Left = 176
      Top = 126
      Width = 6
      Height = 16
      Caption = '/'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label11: TLabel
      Left = 349
      Top = 124
      Width = 13
      Height = 24
      Caption = '='
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -19
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object G: THGrid
      Left = 4
      Top = 153
      Width = 522
      Height = 123
      ColCount = 7
      Ctl3D = True
      DefaultColWidth = 100
      DefaultRowHeight = 18
      Enabled = False
      FixedCols = 0
      RowCount = 6
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goTabs]
      ParentCtl3D = False
      TabOrder = 3
      SortedCol = -1
      Titres.Strings = (
        'Base'
        'Code'
        'Taux'
        'Base factures'
        'Base acompte'
        'Base calcul'#233'e'
        'Ecart')
      Couleur = True
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        50
        38
        45
        89
        89
        93
        85)
    end
    object GFact: THGrid
      Left = 196
      Top = 7
      Width = 330
      Height = 88
      ColCount = 4
      DefaultRowHeight = 16
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goTabs, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 2
      SortedCol = -1
      Titres.Strings = (
        'Num'#233'ro'
        'Paie.'
        'Ech'#233'ance'
        'Montant')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        121
        48
        72
        64)
    end
    object TTC_ACC: THNumEdit
      Tag = 2
      Left = 100
      Top = 54
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object TTC_FACT: THNumEdit
      Tag = 2
      Left = 100
      Top = 22
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object COUV_ACC: THNumEdit
      Tag = 2
      Left = 59
      Top = 124
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      TabOrder = 4
      UseRounding = True
      Validate = False
    end
    object COUV_FACT: THNumEdit
      Tag = 2
      Left = 216
      Top = 124
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      TabOrder = 5
      UseRounding = True
      Validate = False
    end
    object Ratio: THNumEdit
      Left = 401
      Top = 124
      Width = 85
      Height = 21
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntPercentage
      TabOrder = 6
      UseRounding = True
      Validate = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 177
    Top = 216
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Encais 1'
      'Encais 2'
      'Encais 3'
      'Encais 4'
      'D'#233'bit'
      
        '5;?caption?;Validez-vous la conformit'#233' des bases HT de l'#39'acompte' +
        ' avec celles des factures ?;Q;YN;N;N;'
      
        'ATTENTION !  L'#39'acompte est en cours de modification par un autre' +
        ' utilisateur et n'#39'a pas '#233't'#233' valid'#233'. ')
    Left = 135
    Top = 215
  end
end

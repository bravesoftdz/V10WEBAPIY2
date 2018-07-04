object FMsgTVA: TFMsgTVA
  Left = 145
  Top = 162
  Width = 555
  Height = 289
  ActiveControl = BFerme
  BorderIcons = [biSystemMenu]
  Caption = 'Pr'#233'-validation des '#233'ditions de TVA sur encaissements'
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
  object GBFermeDEF: TGroupBox
    Left = 0
    Top = 77
    Width = 547
    Height = 149
    Align = alClient
    Caption = 'ATTENTION !'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object HLabel1: THLabel
      Left = 16
      Top = 22
      Width = 522
      Height = 16
      Caption = 
        'Vous venez de proc'#233'der '#224' l'#39#233'dition li'#233'e au r'#232'glement "indirect" ' +
        'des factures'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel3: THLabel
      Left = 16
      Top = 47
      Width = 387
      Height = 16
      Caption = 'Le traitement de pr'#233'-validation aura pour cons'#233'quence :'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object HLabel4: THLabel
      Left = 35
      Top = 67
      Width = 234
      Height = 16
      Caption = '- de confirmer les montants '#233'dit'#233's'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object HLabel5: THLabel
      Left = 16
      Top = 122
      Width = 333
      Height = 16
      Caption = 'Utilisez "Abandonner" pour ne pas valider ce traitement. '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel6: THLabel
      Left = 35
      Top = 89
      Width = 397
      Height = 16
      Caption = '- de permettre la validation d'#233'finitive des montants '#233'dit'#233's'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 226
    Width = 547
    Height = 36
    Hint = 'Param'#233'trage cl'#244'ture'
    Align = alBottom
    BevelInner = bvLowered
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
    TabOrder = 0
    object BValider: THBitBtn
      Left = 480
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
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
      OnClick = BValiderClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0127_S16G1'
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 513
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Abandonner'
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
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 547
    Height = 77
    Align = alTop
    Caption = 'Rappel des informations de l'#39#233'dition'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object HLabel14: THLabel
      Left = 218
      Top = 21
      Width = 47
      Height = 13
      Caption = 'Edition du'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label1: TLabel
      Left = 398
      Top = 21
      Width = 12
      Height = 13
      Caption = 'au'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel2: THLabel
      Left = 218
      Top = 49
      Width = 49
      Height = 13
      Caption = 'Code TVA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel7: THLabel
      Left = 381
      Top = 49
      Width = 60
      Height = 13
      Caption = 'R'#233'gime TVA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Fdate1: TMaskEdit
      Left = 283
      Top = 17
      Width = 78
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!00/00/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      Text = '  /  /    '
    end
    object FDate2: TMaskEdit
      Left = 451
      Top = 17
      Width = 77
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!00/00/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      Text = '  /  /    '
    end
    object TypeTVA: TRadioGroup
      Left = 9
      Top = 24
      Width = 168
      Height = 41
      Caption = ' TVA '
      Columns = 2
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        '&Collect'#233'e'
        'D'#233'ducti&ble')
      ParentFont = False
      TabOrder = 2
    end
    object FCodeTva: TEdit
      Left = 284
      Top = 45
      Width = 77
      Height = 21
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
    end
    object FRegimeTva: TEdit
      Left = 451
      Top = 45
      Width = 77
      Height = 21
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
    end
  end
  object HMess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Confirmez-vous la pr'#233'validation de l'#39#233'dition des fac' +
        'tures directes?;E;YN;Y;N'
      
        '1;?caption?;Confirmez-vous la pr'#233'validation de l'#39#233'dition des fac' +
        'tures indirectes?;E;YN;Y;N'
      
        '2;?caption?;Confirmez-vous la pr'#233'validation de l'#39#233'dition des aco' +
        'mptes?;E;YN;Y;N'
      
        'Vous venez de proc'#233'der '#224' l'#39#233'dition li'#233'e au r'#232'glement "direct" de' +
        's factures.'
      
        'Vous venez de proc'#233'der '#224' l'#39#233'dition li'#233'e au r'#232'glement "indirect" ' +
        'des factures.'
      
        'Vous venez de proc'#233'der '#224' l'#39#233'dition li'#233'e au r'#232'glement des acompte' +
        's non lettr'#233's.')
    Left = 449
    Top = 121
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 488
    Top = 120
  end
end

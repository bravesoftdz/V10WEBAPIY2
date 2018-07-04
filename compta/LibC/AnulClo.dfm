object FDECLO: TFDECLO
  Left = 292
  Top = 223
  HelpContext = 3245000
  BorderStyle = bsSingle
  Caption = 'Annulation de cl'#244'ture d'#233'finitive'
  ClientHeight = 295
  ClientWidth = 577
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
  object GBFERMEPRO: TGroupBox
    Left = 0
    Top = 77
    Width = 577
    Height = 178
    Align = alClient
    Caption = 'ATTENTION !'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object HLabel7: THLabel
      Left = 20
      Top = 24
      Width = 517
      Height = 20
      AutoSize = False
      Caption = 'Cette op'#233'ration est irr'#233'versible.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel8: THLabel
      Left = 20
      Top = 52
      Width = 491
      Height = 24
      AutoSize = False
      Caption = '1'#176') Elle va annuler la derni'#232're cl'#244'ture provisoire.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel9: THLabel
      Left = 20
      Top = 75
      Width = 510
      Height = 22
      AutoSize = False
      Caption = '2'#176') Elle va d'#233'truire sur l'#39'exercice suivant : '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel11: THLabel
      Left = 20
      Top = 137
      Width = 527
      Height = 22
      AutoSize = False
      Caption = 
        '3'#176') A l'#39'issue de cette op'#233'ration, vous pourrez refaire une cl'#244'tu' +
        're provisoire.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel12: THLabel
      Left = 47
      Top = 96
      Width = 364
      Height = 23
      AutoSize = False
      Caption = '- les '#233'critures d'#39#224'-nouveaux provisoires'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
  end
  object GBFermeDEF: TGroupBox
    Left = 0
    Top = 77
    Width = 577
    Height = 178
    Align = alClient
    Caption = 'ATTENTION !'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object HLabel1: THLabel
      Left = 20
      Top = 24
      Width = 545
      Height = 20
      AutoSize = False
      Caption = 'Cette op'#233'ration est irr'#233'versible.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel2: THLabel
      Left = 20
      Top = 52
      Width = 545
      Height = 24
      AutoSize = False
      Caption = '1'#176') Elle va annuler la derni'#232're cl'#244'ture d'#233'finitive.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel3: THLabel
      Left = 20
      Top = 75
      Width = 550
      Height = 22
      AutoSize = False
      Caption = '2'#176') Elle va d'#233'truire :'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel4: THLabel
      Left = 47
      Top = 95
      Width = 517
      Height = 23
      AutoSize = False
      Caption = '- les '#233'critures de cl'#244'ture sur l'#39'exercice pr'#233'c'#233'dent.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel5: THLabel
      Left = 20
      Top = 137
      Width = 505
      Height = 22
      AutoSize = False
      Caption = 
        '3'#176') L'#39'actuel exercice pr'#233'c'#233'dent (N-1) va devenir l'#39'exercice en c' +
        'ours (N).'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object HLabel6: THLabel
      Left = 47
      Top = 113
      Width = 494
      Height = 23
      AutoSize = False
      Caption = '- les '#233'critures d'#39#224'-nouveaux sur l'#39'exercice en cours.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 255
    Width = 577
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      Hint = 'Param'#233'trage cl'#244'ture'
      ClientHeight = 36
      ClientWidth = 577
      ClientAreaHeight = 36
      ClientAreaWidth = 577
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 542
        Top = 3
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
        Left = 478
        Top = 3
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
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 510
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
        ModalResult = 2
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
      end
    end
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 577
    Height = 77
    Align = alTop
    Caption = 'Dates d'#39'exercice'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    object HLabel10: THLabel
      Left = 292
      Top = 52
      Width = 107
      Height = 13
      AutoSize = False
      Caption = 'Exercice suivant : de'
      FocusControl = DateDebN1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 476
      Top = 52
      Width = 6
      Height = 13
      Caption = #224
      FocusControl = DateFinN1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel13: THLabel
      Left = 12
      Top = 52
      Width = 110
      Height = 13
      AutoSize = False
      Caption = 'Exercice en cours : de'
      FocusControl = DateDebN
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 201
      Top = 52
      Width = 6
      Height = 13
      Caption = #224
      FocusControl = DateFinN
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel14: THLabel
      Left = 12
      Top = 24
      Width = 112
      Height = 13
      AutoSize = False
      Caption = 'Exercice pr'#233'c'#233'dent :de'
      FocusControl = DateDebM
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label1: TLabel
      Left = 202
      Top = 24
      Width = 6
      Height = 13
      Caption = #224
      FocusControl = DateFinM
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HPatienter: THLabel
      Left = 372
      Top = 20
      Width = 177
      Height = 20
      AutoSize = False
      Caption = 'Veuillez patienter'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object Label3: TLabel
      Left = 393
      Top = 52
      Width = 178
      Height = 13
      Caption = 'EXERCICE SUIVANT NON OUVERT'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      Visible = False
    end
    object Label2: TLabel
      Left = 132
      Top = 24
      Width = 201
      Height = 13
      AutoSize = False
      Caption = 'PAS D'#39'EXERCICE PRECEDENT'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      Visible = False
    end
    object DateDebM: TMaskEdit
      Left = 127
      Top = 20
      Width = 65
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
      TabOrder = 4
      Text = '  /  /    '
    end
    object DateDebN1: TMaskEdit
      Left = 399
      Top = 48
      Width = 65
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
    object DateFinN1: TMaskEdit
      Left = 493
      Top = 48
      Width = 65
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
    object DateDebN: TMaskEdit
      Left = 127
      Top = 48
      Width = 65
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
      TabOrder = 2
      Text = '  /  /    '
    end
    object DateFinN: TMaskEdit
      Left = 216
      Top = 48
      Width = 65
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
      TabOrder = 3
      Text = '  /  /    '
    end
    object DateFinM: TMaskEdit
      Left = 216
      Top = 20
      Width = 65
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
      TabOrder = 5
      Text = '  /  /    '
    end
  end
  object Confirmation: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Annulation de cl'#244'ture;1'#176' confirmation : confirmez-vous l'#39'annul' +
        'ation de la cl'#244'ture?;E;YN;Y;N'
      
        '1;Annulation de cl'#244'ture;2'#176' confirmation : confirmez-vous l'#39'annul' +
        'ation de la cl'#244'ture?;E;YN;Y;N'
      
        '2;Annulation de cl'#244'ture;Aucune cl'#244'ture comptable n'#39'a '#233't'#233' faite s' +
        'ur le produit;E;O;O;O'
      
        '3;Annulation de cl'#244'ture;Veuillez vous positionner sur l'#39'en-cours' +
        ' (date d'#39'entr'#233'e);E;O;O;O'
      '4;;;E;0;0;0;'
      'ATTENTION ! Probl'#232'me pendant l'#39'annulation de cl'#244'ture'
      'Annulation de cl'#244'ture provisoire'
      'Annulation de cl'#244'ture d'#233'finitive'
      'Annulation de cl'#244'ture'
      
        'Attention : Cette op'#233'ration n'#39'est accessible qu'#39#224' l'#39'administrate' +
        'ur.'
      'Veuillez contacter votre administrateur.;E;O;O;O;'
      '11;NE PAS TOUCHER !!'
      
        '12;Annulation de cl'#244'ture;Attention, apr'#232's avoir recl'#244'tur'#233' cet ex' +
        'ercice, il est imp'#233'ratif de r'#233'ouvrir manuellement l'#39'exercice en ' +
        'cours.;E;O;O;O'
      
        '13;Pour la conformit'#233' stricte avec la norme NF 203 (et le BOI du' +
        ' 24/01/2006) cette fonction n'#39'est plus disponible'
      ''
      ' '
      ' ')
    Left = 409
    Top = 93
  end
  object Q: TQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    Left = 340
    Top = 172
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 464
    Top = 92
  end
end

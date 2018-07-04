object F_Paie_cjdc: TF_Paie_cjdc
  Left = 317
  Top = 209
  BorderIcons = []
  BorderStyle = bsDialog
  Caption = ' Service Web Jedeclare'
  ClientHeight = 404
  ClientWidth = 538
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnActivate = ActiverFiche
  PixelsPerInch = 96
  TextHeight = 13
  object TPanel
    Left = 1
    Top = 358
    Width = 532
    Height = 49
    TabOrder = 1
    object Valider: TBitBtn
      Left = 439
      Top = 5
      Width = 33
      Height = 33
      Hint = 'Valider'
      Default = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = ValiderClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object Annuler: TBitBtn
      Left = 482
      Top = 4
      Width = 32
      Height = 33
      Hint = 'Annuler'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Visible = False
      OnClick = AnnulerClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333000033338833333333333333333F333333333333
        0000333911833333983333333388F333333F3333000033391118333911833333
        38F38F333F88F33300003339111183911118333338F338F3F8338F3300003333
        911118111118333338F3338F833338F3000033333911111111833333338F3338
        3333F8330000333333911111183333333338F333333F83330000333333311111
        8333333333338F3333383333000033333339111183333333333338F333833333
        00003333339111118333333333333833338F3333000033333911181118333333
        33338333338F333300003333911183911183333333383338F338F33300003333
        9118333911183333338F33838F338F33000033333913333391113333338FF833
        38F338F300003333333333333919333333388333338FFF830000333333333333
        3333333333333333333888330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
  end
  object Panel1: TPanel
    Left = 6
    Top = 4
    Width = 527
    Height = 353
    TabOrder = 0
    object Chemin: TLabel
      Left = 108
      Top = 114
      Width = 3
      Height = 13
    end
    object Label1: TLabel
      Left = 11
      Top = 115
      Width = 59
      Height = 13
      Caption = 'Traitement  :'
    end
    object GroupBox1: TGroupBox
      Left = 10
      Top = 4
      Width = 342
      Height = 37
      TabOrder = 1
      object Label2: TLabel
        Left = 12
        Top = 16
        Width = 42
        Height = 13
        Caption = 'Emetteur'
      end
      object ComboEmetteurs: THValComboBox
        Left = 78
        Top = 12
        Width = 255
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        Text = 'ComboEmetteurs'
        TagDispatch = 0
      end
    end
    object ReceptionBarre: TProgressBar
      Left = 11
      Top = 49
      Width = 340
      Height = 18
      TabOrder = 0
    end
    object RichEdit: TRichEdit
      Left = 12
      Top = 88
      Width = 505
      Height = 254
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'Times New Roman'
      Font.Style = []
      Lines.Strings = (
        'RichEdit')
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 2
    end
  end
end

object FAudit: TFAudit
  Left = 356
  Top = 225
  HelpContext = 7745100
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Audit avant cl'#244'ture'
  ClientHeight = 246
  ClientWidth = 427
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
  object ToolbarButton971: TToolbarButton97
    Left = 331
    Top = 2
    Width = 28
    Height = 27
    Hint = 'Acc'#233'der '#224' la cl'#244'ture d'#233'finitive'
    Default = True
    DisplayMode = dmGlyphOnly
    Caption = 'Cl'#244'ture'
    Flat = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 344
    Height = 208
    Align = alClient
    TabOrder = 0
    object V1: THLabel
      Left = 8
      Top = 24
      Width = 104
      Height = 13
      Caption = 'V'#233'rification comptable'
    end
    object V2: THLabel
      Left = 10
      Top = 59
      Width = 157
      Height = 13
      Caption = 'V'#233'rification des p'#233'riodes valid'#233'es'
    end
    object V4: THLabel
      Left = 10
      Top = 97
      Width = 206
      Height = 13
      Caption = 'V'#233'rification du solde de comptes particuliers'
    end
    object V6: THLabel
      Left = 10
      Top = 174
      Width = 87
      Height = 13
      Caption = 'Contr'#244'le de caisse'
    end
    object V5: THLabel
      Left = 10
      Top = 136
      Width = 281
      Height = 13
      Caption = 'V'#233'rification des '#233'critures de simu., r'#233'vision ou abonnements'
    end
    object B1: TToolbarButton97
      Left = 308
      Top = 17
      Width = 28
      Height = 27
      Hint = 'V'#233'rification comptable'
      Default = True
      DisplayMode = dmGlyphOnly
      Caption = 'Cl'#244'ture'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = B1Click
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object B2: TToolbarButton97
      Left = 308
      Top = 52
      Width = 28
      Height = 27
      Hint = 'V'#233'rification des p'#233'riodes valid'#233'es'
      Default = True
      DisplayMode = dmGlyphOnly
      Caption = 'Cl'#244'ture'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = B2Click
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object B4: TToolbarButton97
      Left = 308
      Top = 90
      Width = 28
      Height = 27
      Hint = 'V'#233'rification du solde de comptes particuliers'
      Default = True
      DisplayMode = dmGlyphOnly
      Caption = 'Cl'#244'ture'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = B4Click
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object B5: TToolbarButton97
      Left = 308
      Top = 129
      Width = 28
      Height = 27
      Hint = 'V'#233'rif. des '#233'critures particuli'#232'res'
      Default = True
      DisplayMode = dmGlyphOnly
      Caption = 'Cl'#244'ture'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = B5Click
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object B6: TToolbarButton97
      Left = 308
      Top = 167
      Width = 28
      Height = 27
      Hint = 'Contr'#244'le de caisse'
      Default = True
      DisplayMode = dmGlyphOnly
      Caption = 'Cl'#244'ture'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = B6Click
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 208
    Width = 427
    Height = 38
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 427
      Caption = ' '
      ClientAreaHeight = 34
      ClientAreaWidth = 427
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 393
        Top = 2
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
        Left = 331
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Acc'#233'der '#224' la cl'#244'ture d'#233'finitive'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Cl'#244'ture'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
        Left = 362
        Top = 2
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
  object Panel2: TPanel
    Left = 344
    Top = 0
    Width = 83
    Height = 208
    Align = alRight
    TabOrder = 2
    object R1: THLabel
      Left = 19
      Top = 20
      Width = 39
      Height = 13
      Caption = 'R'#233'sultat'
    end
    object R2: THLabel
      Left = 19
      Top = 60
      Width = 39
      Height = 13
      Caption = 'R'#233'sultat'
    end
    object R4: THLabel
      Left = 19
      Top = 99
      Width = 39
      Height = 13
      Caption = 'R'#233'sultat'
    end
    object R5: THLabel
      Left = 19
      Top = 139
      Width = 39
      Height = 13
      Caption = 'R'#233'sultat'
    end
    object R6: THLabel
      Left = 19
      Top = 178
      Width = 39
      Height = 13
      Caption = 'R'#233'sultat'
    end
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Les p'#233'riodes ne sont pas valid'#233'es, d'#233'sirez-vous les ' +
        'valider ?;Q;YN;N;N;'
      
        '1;?caption?;Toutes les '#233'ditions l'#233'gales ne sont pas '#233'dit'#233'es, d'#233's' +
        'irez-vous les '#233'diter ?;Q;YN;N;N;'
      '2;?caption?;Toutes les p'#233'riodes sont valid'#233'es.;A;O;O;O;'
      
        '3;?caption?;Toutes les '#233'ditions l'#233'gales ont '#233't'#233' effectu'#233'es ;A;O;' +
        'O;O;'
      '4;?caption?;;E;O;O;O;')
    Left = 192
    Top = 20
  end
  object MsgLibel: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'A faire'
      'Correct'
      'Incorrect'
      'V'#233'rification des '#233'critures de simulation ou abonnements')
    Left = 220
    Top = 88
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 252
    Top = 24
  end
end

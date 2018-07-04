object FAffBqe: TFAffBqe
  Left = 101
  Top = 184
  HelpContext = 7503700
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Affectation pr'#233'visionnelle des banques'
  ClientHeight = 339
  ClientWidth = 711
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 305
    Width = 711
    Height = 34
    Align = alBottom
    BevelOuter = bvLowered
    TabOrder = 0
    object BImprimer: THBitBtn
      Left = 570
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BValider: THBitBtn
      Left = 602
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BValiderClick
      GlobalIndexImage = 'Z0003_S16G2'
    end
    object BAnnuler: THBitBtn
      Left = 634
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BAnnulerClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 666
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object bZoomPiece: THBitBtn
      Left = 538
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Voir l'#39#233'criture'
      HelpContext = 1710
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
      OnClick = bZoomPieceClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0016_S16G1'
      IsControl = True
    end
  end
  object GBqe: THGrid
    Left = 0
    Top = 184
    Width = 711
    Height = 121
    Align = alClient
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    TabOrder = 1
    SortedCol = -1
    Titres.Strings = (
      'Compte'
      'Libell'#233
      'Solde'
      'Montant affect'#233
      'Solde pr'#233'visionnel')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      90
      155
      100
      98
      102)
  end
  object Split: TSplitControl
    Left = 0
    Top = 180
    Width = 711
    Height = 4
    Cursor = crVSplit
    Align = alTop
    Caption = ' '
    TabOrder = 2
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 711
    Height = 180
    Align = alTop
    Caption = ' '
    TabOrder = 3
    object G: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 709
      Height = 148
      Align = alTop
      ColCount = 14
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 0
      OnDblClick = GDblClick
      OnMouseUp = GMouseUp
      SortedCol = -1
      ListeParam = 'ENCDEC'
      Couleur = True
      MultiSelect = True
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
    end
    object Panel3: TPanel
      Left = 1
      Top = 149
      Width = 709
      Height = 33
      Align = alTop
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 1
      object H_TOTDEVISE: TLabel
        Left = 412
        Top = 9
        Width = 24
        Height = 13
        Caption = 'Total'
      end
      object bUp: THBitBtn
        Left = 35
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'saffecter les '#233'ch'#233'ances s'#233'lectionn'#233'es'
        HelpContext = 1710
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = bUpClick
        Margin = 5
        Spacing = -1
        GlobalIndexImage = 'Z0118_S16G1'
        IsControl = True
      end
      object bDown: THBitBtn
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Affecter les '#233'ch'#233'ances s'#233'lectionn'#233'es '#224' la banque s'#233'lectionn'#233'e'
        HelpContext = 1710
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = bDownClick
        Margin = 5
        Spacing = -1
        GlobalIndexImage = 'Z0192_S16G1'
        IsControl = True
      end
      object ER_SOLDED: THNumEdit
        Tag = 2
        Left = 444
        Top = 5
        Width = 126
        Height = 21
        Color = clBtnFace
        Ctl3D = True
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        NumericType = ntDC
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        UseRounding = True
        Validate = False
      end
      object cVisu: TCheckBox
        Left = 259
        Top = 6
        Width = 51
        Height = 17
        Caption = 'Visu'
        Color = clYellow
        ParentColor = False
        TabOrder = 4
        Visible = False
      end
      object Devise: TEdit
        Left = 147
        Top = 5
        Width = 90
        Height = 21
        Color = clYellow
        TabOrder = 5
        Visible = False
      end
      object bSelectAll: THBitBtn
        Left = 67
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lect./Tout d'#233'select.'
        HelpContext = 1710
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = bSelectAllClick
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000130B0000130B00001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          333333333333333333FF33333333333330003333333333333888333333333333
          300033FFFFFF3333388839999993333333333888888F3333333F399999933333
          3300388888833333338833333333333333003333333333333388333333333333
          3333333333333333333F3333333E3333330033333F33333333883333333EC333
          330033338F3333333388EEEEEEEECCE3333333F88FFFFFFF3FF3ECCCCCCCCCCC
          399338888888888F88F33CCCCCCCCCCC399338888888888388333333333ECC33
          333333388F33333333FF33333333C33330003333833333333888333333333333
          3000333333333333388833333333333333333333333333333333}
        Margin = 3
        NumGlyphs = 2
        Spacing = -1
        IsControl = True
      end
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Attention : validation non effectu'#233'e !'
      
        '1;?caption?;Aucune '#233'ch'#233'ance n'#39'est s'#233'lectionn'#233'e, op'#233'ration imposs' +
        'ible.;W;O;O;O;'
      
        '2;?caption?;Attention : certaines '#233'ch'#233'ances n'#39'ont pas '#233't'#233' affect' +
        #233'es, confirmer la validation ?;Q;YN;Y;N;'
      '')
    Left = 197
    Top = 76
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 152
    Top = 74
  end
end

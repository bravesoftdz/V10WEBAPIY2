inherited FCpyBudMu: TFCpyBudMu
  Left = 216
  Top = 221
  Width = 606
  Height = 382
  HelpContext = 15271000
  Caption = 'Copie multiple de budget'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object PG: TPanel [0]
    Left = 0
    Top = 0
    Width = 353
    Height = 316
    Align = alClient
    TabOrder = 1
    object Bevel1: TBevel
      Left = 1
      Top = 1
      Width = 351
      Height = 30
      Align = alTop
    end
    object TNatbud: TLabel
      Left = 12
      Top = 9
      Width = 95
      Height = 13
      Caption = 'Budgets sources'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Nb1: TLabel
      Left = 174
      Top = 9
      Width = 8
      Height = 13
      Caption = '0'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Tex1: TLabel
      Left = 202
      Top = 9
      Width = 132
      Height = 13
      Caption = 'ligne(s) sélectionnée(s)'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object FListe: THGrid
      Left = 1
      Top = 31
      Width = 351
      Height = 284
      Align = alClient
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      OnDblClick = FListeDblClick
      OnKeyDown = FListeKeyDown
      OnMouseDown = FListeMouseDown
      SortedCol = -1
      Titres.Strings = (
        'Code;G'
        'Libellé;G'
        'Période;C'
        'Nb Piece;C')
      Couleur = True
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        40
        174
        55
        55
        0)
    end
    object PerdebS: THValComboBox
      Left = 50
      Top = 113
      Width = 25
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 1
      Visible = False
      TagDispatch = 0
    end
    object PerfinS: THValComboBox
      Left = 46
      Top = 141
      Width = 25
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 2
      Visible = False
      TagDispatch = 0
    end
  end
  inherited Dock971: TDock97
    Top = 316
    Width = 598
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 598
      ClientAreaHeight = 35
      ClientAreaWidth = 598
      inherited BValider: TToolbarButton97
        Left = 502
        Top = 4
      end
      inherited BFerme: TToolbarButton97
        Left = 534
        Top = 4
      end
      inherited HelpBtn: TToolbarButton97
        Left = 566
        Top = 4
      end
      inherited bDefaire: TToolbarButton97
        Top = 4
      end
      inherited Binsert: TToolbarButton97
        Top = 4
      end
      inherited BDelete: TToolbarButton97
        Top = 4
      end
      inherited BImprimer: TToolbarButton97
        Left = 470
        Top = 4
      end
      object BZoom: TToolbarButton97
        Left = 133
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Vision consolidée du budget'
        DisplayMode = dmGlyphOnly
        Caption = 'Zoom'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC00000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
          DDDDD0000000D00DDDDDDDDDDDDDD0000000D000DDDDDDDDDDDDD0000000DD00
          0DDDDDDDDDDDD0000000DDD000D800008DDDD0000000DDDD0000777700DDD000
          0000DDDDD08EE777780DD0000000DDDD807E77777708D0000000DDDD07E77777
          7770D0000000DDDD077777777770D0000000DDDD077777777E70D0000000DDDD
          077777777E70D0000000DDDD80777777EE08D0000000DDDDD08777EEE80DD000
          0000DDDDDD00777700DDD0000000DDDDDDD800008DDDD0000000DDDDDDDDDDDD
          DDDDD0000000}
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BZoomClick
        IsControl = True
      end
      object BOpt: TToolbarButton97
        Left = 101
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Paramètrage des natures et du coefficient de copie par budget'
        DisplayMode = dmGlyphOnly
        Caption = 'Paramètres'
        Flat = False
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC000000CE0E0000D80E00001000000000000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        ParentShowHint = False
        ShowHint = True
        OnClick = BOptClick
      end
    end
  end
  object PD: TPanel [2]
    Left = 353
    Top = 0
    Width = 245
    Height = 316
    Align = alRight
    TabOrder = 2
    object Bevel2: TBevel
      Left = 1
      Top = 1
      Width = 243
      Height = 30
      Align = alTop
    end
    object TBudJal: TLabel
      Left = 8
      Top = 43
      Width = 34
      Height = 13
      Caption = '&Budget'
    end
    object BudD: TLabel
      Left = 69
      Top = 9
      Width = 107
      Height = 13
      Caption = 'Budget destination'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object BudJal: THValComboBox
      Left = 55
      Top = 39
      Width = 179
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = BudJalChange
      TagDispatch = 0
    end
    object GbComplement: TGroupBox
      Left = 8
      Top = 155
      Width = 227
      Height = 119
      Caption = ' Informations complémentaires '
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      object TExoDebD: TLabel
        Left = 5
        Top = 21
        Width = 29
        Height = 13
        Caption = 'Début'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object TExoFinD: TLabel
        Left = 5
        Top = 45
        Width = 14
        Height = 13
        Caption = 'Fin'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object TPerDebD: TLabel
        Left = 5
        Top = 69
        Width = 29
        Height = 13
        Caption = 'Début'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object TPerFinD: TLabel
        Left = 5
        Top = 93
        Width = 14
        Height = 13
        Caption = 'Fin'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object ExoDebD: THValComboBox
        Left = 42
        Top = 17
        Width = 177
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object ExoFinD: THValComboBox
        Left = 42
        Top = 41
        Width = 177
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object PerDebD: THValComboBox
        Left = 42
        Top = 64
        Width = 177
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 2
        TagDispatch = 0
      end
      object PerFinD: THValComboBox
        Left = 42
        Top = 89
        Width = 177
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 3
        TagDispatch = 0
      end
    end
    object RgCopi: TRadioGroup
      Left = 8
      Top = 69
      Width = 227
      Height = 83
      Caption = ' Méthode de copie de pièce '
      ItemIndex = 0
      Items.Strings = (
        '&Vider le budget destination'
        '&Duplication'
        '&Consolidation')
      TabOrder = 1
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ligne sélectionnée'
      'lignes sélectionnées'
      
        '2;Copie multiple de budget;Aucun budget n'#39'a été sélectionné pour' +
        ' la copie.;W;O;O;O;'
      
        '3;Copie multiple de budget;Aucun budget sélectionné ne pourra êt' +
        're copié. Caractéristiques différentes ou aucune pièce à copier.' +
        ';W;O;O;O;'
      
        '4;Copie multiple de budget;Désirez-vous copier les budgets sélec' +
        'tionnés?;Q;YN;N;N;'
      
        '5;Copie multiple de budget;Désirez-vous visualiser les écritures' +
        ' budgétaires?;Q;YN;N;N;'
      'ATTENTION ! Recopie non effectuée.'
      
        '7;Copie multiple de budget;La recopie de budget s'#39'est correcteme' +
        'nt effectuée !;E;O;O;O;')
    Left = 265
    Top = 149
  end
  object QNbEcr: TQuery
    DatabaseName = 'SOC'
    SQL.Strings = (
      'Select BE_NUMEROPIECE,BE_NATUREBUD From BUDECR '
      'Where BE_BUDJAL=:BudJ Group by BE_NUMEROPIECE,BE_NATUREBUD')
    Left = 156
    Top = 113
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'BudJ'
        ParamType = ptUnknown
      end>
  end
  object QInser: TQuery
    DatabaseName = 'SOC'
    UpdateMode = upWhereKeyOnly
    Left = 196
    Top = 216
  end
end

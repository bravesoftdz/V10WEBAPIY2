inherited FUExport: TFUExport
  Left = 196
  Top = 224
  Width = 750
  Height = 484
  HelpContext = 1300
  Caption = 'G'#233'n'#233'ration simple'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 415
    Width = 742
    inherited PBouton: TToolWindow97
      ClientWidth = 742
      ClientAreaWidth = 742
      inherited BValider: TToolbarButton97
        Left = 643
      end
      inherited BFerme: TToolbarButton97
        Left = 675
      end
      inherited HelpBtn: TToolbarButton97
        Tag = 1300
        Left = 707
      end
      inherited bDefaire: TToolbarButton97
        Hint = 'Vider la liste'
        Caption = 'Vider la liste'
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 611
      end
    end
  end
  object PageControl1: TPageControl [1]
    Left = 0
    Top = 0
    Width = 742
    Height = 415
    Align = alClient
    TabOrder = 1
  end
  object Panel1: TPanel [2]
    Left = 0
    Top = 0
    Width = 742
    Height = 415
    Align = alClient
    Caption = 'Panel1'
    TabOrder = 2
    object GD: THGrid
      Left = 1
      Top = 83
      Width = 740
      Height = 331
      Align = alClient
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
      ParentFont = False
      TabOrder = 0
      OnDblClick = GDDblClick
      OnSelectCell = GDSelectCell
      SortedCol = -1
      Titres.Strings = (
        'Fichier d'#39'entr'#233'e'
        'Script'
        'Fichier de sortie'
        'Effectu'#233
        'Fin de traitement')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        238
        164
        116
        115
        99)
    end
    object HPanel1: THPanel
      Left = 1
      Top = 1
      Width = 740
      Height = 82
      Align = alTop
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      DesignSize = (
        740
        82)
      object LTable: TLabel
        Left = 9
        Top = 9
        Width = 82
        Height = 13
        Caption = 'Script '#224' appliquer'
      end
      object TFListefichiers: TLabel
        Left = 324
        Top = 9
        Width = 77
        Height = 13
        Anchors = [akTop, akRight]
        Caption = 'Fichiers d'#39'entr'#233'e'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BValide: TToolbarButton97
        Left = 707
        Top = 5
        Width = 27
        Height = 21
        Hint = 'Valider les fichiers'
        Alignment = taRightJustify
        Anchors = [akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir'
        NoBorder = True
        ParentShowHint = False
        ShowHint = True
        OnClick = BValideClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object LFichier: TLabel
        Left = 9
        Top = 34
        Width = 76
        Height = 13
        Caption = 'Fichier de Sortie'
      end
      object Label12: TLabel
        Left = 9
        Top = 61
        Width = 78
        Height = 13
        Caption = 'Correspondance'
      end
      object Table: THCritMaskEdit
        Left = 106
        Top = 5
        Width = 157
        Height = 21
        MaxLength = 12
        TabOrder = 0
        OnClick = TableClick
        OnExit = TableExit
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = TableElipsisClick
      end
      object FListefichiers: THCritMaskEdit
        Left = 414
        Top = 5
        Width = 282
        Height = 21
        Anchors = [akTop, akRight]
        TabOrder = 1
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = FListefichiersElipsisClick
      end
      object FILENAME: THCritMaskEdit
        Left = 106
        Top = 30
        Width = 296
        Height = 21
        TabOrder = 2
        TagDispatch = 0
        DataType = 'SAVEFILE(*.TRA;*.ZIP;*.TRT;*.*)'
        ElipsisButton = True
      end
      object Memefichier: TCheckBox
        Left = 414
        Top = 32
        Width = 235
        Height = 17
        Caption = 'Int'#233'grer dans le m'#234'me fichier de sortie'
        TabOrder = 3
      end
      object Profile: THCritMaskEdit
        Left = 106
        Top = 57
        Width = 296
        Height = 21
        MaxLength = 35
        TabOrder = 4
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = ProfileElipsisClick
      end
      object BCompress: TCheckBox
        Left = 414
        Top = 59
        Width = 84
        Height = 17
        Caption = 'Compresser'
        TabOrder = 5
      end
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 649
    Top = 52
  end
end

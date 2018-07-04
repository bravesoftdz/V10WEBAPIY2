object FTarifRapide: TFTarifRapide
  Left = 162
  Top = 213
  Width = 434
  Height = 363
  HelpContext = 110000065
  AutoSize = True
  BorderIcons = [biSystemMenu]
  Caption = 'Saisie Rapide Tarif'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PENTETE: THPanel
    Left = 0
    Top = 0
    Width = 418
    Height = 97
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_ARTICLE: THLabel
      Left = 4
      Top = 12
      Width = 57
      Height = 13
      Caption = 'Code Article'
      FocusControl = GF_ARTICLE
    end
    object TGF_DEPOT: THLabel
      Left = 4
      Top = 38
      Width = 57
      Height = 13
      Caption = 'Code D'#233'p'#244't'
      FocusControl = GF_DEPOT
    end
    object TGF_DATEDEBUT: THLabel
      Left = 144
      Top = 38
      Width = 55
      Height = 13
      Caption = 'Date D'#233'but'
      FocusControl = GF_DATEDEBUT
    end
    object TGF_DATEFIN: THLabel
      Left = 288
      Top = 38
      Width = 40
      Height = 13
      Caption = 'Date Fin'
    end
    object TGF_LIBELLE: THLabel
      Left = 4
      Top = 65
      Width = 31
      Height = 13
      Caption = 'Intitul'#233
      FocusControl = GF_LIBELLE
    end
    object TGF_LIBELLEARTICLE: THLabel
      Left = 168
      Top = 12
      Width = 237
      Height = 13
      AutoSize = False
      FocusControl = GF_ARTICLE
    end
    object TGF_TARIFTIERS: THLabel
      Left = 204
      Top = 65
      Width = 47
      Height = 13
      Caption = 'Tarif Tiers'
      FocusControl = GF_TARIFTIERS
    end
    object GF_ARTICLE: THCritMaskEdit
      Left = 64
      Top = 8
      Width = 93
      Height = 21
      Color = clActiveBorder
      Enabled = False
      ReadOnly = True
      TabOrder = 0
      TagDispatch = 0
    end
    object GF_DATEDEBUT: THCritMaskEdit
      Left = 204
      Top = 34
      Width = 81
      Height = 21
      EditMask = '!99 >L<LL 0000;1;_'
      MaxLength = 11
      TabOrder = 2
      Text = '           '
      OnExit = GF_DATEDEBUTExit
      TagDispatch = 0
      OpeType = otDate
      DefaultDate = od1900
      ElipsisButton = True
      OnElipsisClick = GF_DATEDEBUTElipsisClick
      ControlerDate = True
    end
    object GF_DATEFIN: THCritMaskEdit
      Left = 332
      Top = 34
      Width = 77
      Height = 21
      EditMask = '!99 >L<LL 0000;1;_'
      MaxLength = 11
      TabOrder = 3
      Text = '           '
      OnExit = GF_DATEFINExit
      TagDispatch = 0
      OpeType = otDate
      DefaultDate = od2099
      ElipsisButton = True
      OnElipsisClick = GF_DATEFINElipsisClick
      ControlerDate = True
    end
    object GF_LIBELLE: THCritMaskEdit
      Left = 64
      Top = 61
      Width = 129
      Height = 21
      TabOrder = 4
      OnExit = GF_LIBELLEExit
      TagDispatch = 0
    end
    object GF_DEPOT: THValComboBox
      Left = 64
      Top = 34
      Width = 73
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      TagDispatch = 0
      DataType = 'GCDEPOT'
    end
    object GF_TARIFTIERS: THValComboBox
      Left = 260
      Top = 61
      Width = 149
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 5
      OnExit = GF_TARIFTIERSExit
      TagDispatch = 0
      DataType = 'TTTARIFCLIENT'
    end
  end
  object G_SaisRap: THGrid
    Tag = 1
    Left = 0
    Top = 100
    Width = 418
    Height = 161
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goTabs]
    TabOrder = 3
    OnEnter = G_SaisRapEnter
    OnExit = G_SaisRapExit
    SortedCol = 0
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = G_SaisRapRowEnter
    OnRowExit = G_SaisRapRowExit
    OnCellEnter = G_SaisRapCellEnter
    OnCellExit = G_SaisRapCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
  end
  object PBOUTONS: TDock97
    Left = 0
    Top = 291
    Width = 426
    Height = 38
    AllowDrag = False
    FixAlign = True
    Position = dpBottom
    object PBARREOUTIL: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 426
      Caption = 'PBARREOUTIL'
      ClientAreaHeight = 34
      ClientAreaWidth = 426
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 335
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Left = 367
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
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 303
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer les tarifs'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Enregistrer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
        Layout = blGlyphTop
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BCondAplli: TToolbarButton97
        Left = 7
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Conditions d'#39'applications tarifaires'
        DisplayMode = dmGlyphOnly
        Caption = 'Modifier'
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BCondAplliClick
        GlobalIndexImage = 'Z0105_S16G2'
      end
    end
  end
  object PPIED: THPanel
    Left = 1
    Top = 260
    Width = 417
    Height = 37
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_REMISE: THLabel
      Left = 8
      Top = 12
      Width = 84
      Height = 13
      Caption = 'Remise r'#233'sultante'
      FocusControl = GF_REMISE
    end
    object GF_REMISE: THCritMaskEdit
      Left = 105
      Top = 8
      Width = 125
      Height = 21
      Color = clActiveBorder
      Enabled = False
      ReadOnly = True
      TabOrder = 0
      OnChange = GF_REMISEChange
      TagDispatch = 0
      OpeType = otReel
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 108
    Top = 192
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;L'#39'intitul'#233' est obligatoire;W;O;O;O;'
      '1;?caption?;La date de d'#233'but <= La date de fin;W;O;O;O; '
      
        '2;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      '3;?caption?;Les bornes de quantit'#233's sont incoh'#233'rentes.;W;O;O;O;'
      '4;?caption?;Le Tarif Tiers est obligatoire;W;O;O;O;'
      '5;?caption?;La date de d'#233'but est obligatoire;W;O;O;O;'
      '6;?caption?;La date de fin est obligatoire;W;O;O;O;'
      '7;?caption?;L'#39'un;W;O;O;O;')
    Left = 160
    Top = 196
  end
end

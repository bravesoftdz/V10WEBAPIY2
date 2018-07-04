object FDimension: TFDimension
  Left = 20
  Top = 180
  Width = 649
  Height = 348
  HelpContext = 110000045
  Caption = 'Param'#233'trage des dimensions'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 268
    Width = 633
    Height = 41
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 37
      ClientWidth = 633
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 37
      ClientAreaWidth = 633
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object bTriDesc2: TToolbarButton97
        Left = 572
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tri croissant'
        Caption = 'Tri d'#233'croissant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = TriAscClick
        GlobalIndexImage = 'Z0319_S16G1'
      end
      object bTriAsc: TToolbarButton97
        Left = 541
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tri d'#233'croissant'
        Caption = 'Tri croissant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = TriDescClick
        GlobalIndexImage = 'Z0818_S16G1'
      end
      object bTriDesc: TToolbarButton97
        Left = 572
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tri croissant'
        Caption = 'Tri d'#233'croissant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = TriAscClick
        GlobalIndexImage = 'Z1206_S16G1'
      end
      object bTriAsc2: TToolbarButton97
        Left = 541
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tri d'#233'croissant'
        Caption = 'Tri croissant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = TriDescClick
        GlobalIndexImage = 'Z0171_S16G1'
      end
      object BDelLigne2: TToolbarButton97
        Left = 395
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer une taille'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BDelLigneClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BDecaleApresClick2: TToolbarButton97
        Left = 509
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'placer apr'#232's'
        Caption = 'D'#233'placer apr'#232's'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDecaleApres
        GlobalIndexImage = 'Z0397_S16G1'
      end
      object BDecaleApresClick: TToolbarButton97
        Left = 509
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'placer '#224' droite'
        Caption = 'D'#233'placer '#224' droite'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDecaleApres
        GlobalIndexImage = 'Z0497_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 61
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
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
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 93
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Default = True
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
        Left = 124
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 157
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
          A600000000000000000000000000000000000000000000000000000000000000
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
          03030606030303030303030303030303030303FFFF0303030303030303030303
          0303030303060404060303030303030303030303030303F8F8FF030303030303
          030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
          0303030303030303030303030303FE060603030303030303030303030303F8FF
          FFF8FF0303030303030303030303030303030303030303030303030303030303
          030303F8F8030303030303030303030303030303030304040603030303030303
          0303030303030303FFFF03030303030303030303030303030306060604030303
          0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
          0403030303030303030303030303F8FF03F8FF03030303030303030303030303
          03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
          030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
          0303030303030303030303FE060606040303030303030303030303F8FF0303F8
          FF030303030303030303030404030303FE060606040303030303030303FF0303
          F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
          F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
          030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
          06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
          060606060303030303030303F8FFFF030303030303F803030303030303030303
          0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
          0303030303030303030303030303030303030303030303F8F8F8F8F803030303
          0303}
        GlyphMask.Data = {00000000}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = HelpBtnClick
        IsControl = True
      end
      object BInsLigne: TToolbarButton97
        Left = 364
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une taille'
        Caption = 'Ins'#233'rer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsLigneClick
        GlobalIndexImage = 'Z0999_S16G0'
      end
      object BDelLigne: TToolbarButton97
        Left = 395
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer une taille'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BDelLigneClick
        GlobalIndexImage = 'Z0375_S16G1'
      end
      object bColonne: TToolbarButton97
        Left = 208
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Afficher en colonnes'
        GroupIndex = 100
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bColonneClick
        GlobalIndexImage = 'Z0519_S16G1'
      end
      object bLigne: TToolbarButton97
        Left = 236
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Afficher en lignes'
        GroupIndex = 100
        Down = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bColonneClick
        GlobalIndexImage = 'Z0291_S16G1'
      end
      object TAjoutDim: TToolbarButton97
        Left = 296
        Top = 4
        Width = 37
        Height = 25
        Hint = 'Ajouter une dimension'
        ParentShowHint = False
        ShowHint = True
        OnClick = TAjoutDimClick
        GlobalIndexImage = 'Z0202_S16G1'
      end
      object Recherche: TToolbarButton97
        Left = 447
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher une taille'
        Caption = 'Recherche'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = RechercheClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BDecaleAvantClick: TToolbarButton97
        Left = 478
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'placer '#224' gauche'
        Caption = 'D'#233'placer '#224' gauche'
        DisplayMode = dmGlyphOnly
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BDecaleAvant
        GlobalIndexImage = 'Z0077_S16G2'
      end
      object BDecaleAvantClick2: TToolbarButton97
        Left = 478
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'placer avant'
        Caption = 'D'#233'placer avant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDecaleAvant
        GlobalIndexImage = 'Z0415_S16G1'
      end
      object BInsLigne2: TToolbarButton97
        Left = 364
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une taille'
        Caption = 'Ins'#233'rer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsLigneClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object FAutoSave: TCheckBox
        Left = 9
        Top = 8
        Width = 44
        Height = 17
        Caption = 'Auto'
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 0
        Visible = False
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 41
    Width = 633
    Height = 227
    Align = alClient
    DefaultColWidth = 50
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
    TabOrder = 1
    OnColumnMoved = FListeDimMoved
    OnRowMoved = FListeDimMoved
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Grille\Taille'
      '1'
      '2'
      '3'
      '4'
      '5'
      '6')
    Couleur = False
    MultiSelect = True
    TitleBold = False
    TitleCenter = True
    OnCellEnter = FListeCellEnter
    OnCellExit = FListeCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    ColWidths = (
      50
      129
      50
      50
      50)
    RowHeights = (
      18
      17
      19
      18
      18)
  end
  object FGrilles: THValComboBox
    Left = 208
    Top = 180
    Width = 97
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 2
    Visible = False
    TagDispatch = 0
  end
  object FGrillesLibres: THValComboBox
    Left = 208
    Top = 204
    Width = 97
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    Sorted = True
    TabOrder = 3
    Visible = False
    TagDispatch = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 633
    Height = 41
    Align = alTop
    Caption = ' '
    TabOrder = 4
    object HLabel1: THLabel
      Left = 32
      Top = 12
      Width = 94
      Height = 13
      Caption = '&Type de dimensions'
      FocusControl = FTypeDim
    end
    object FTypeDim: THValComboBox
      Left = 136
      Top = 8
      Width = 213
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FTypeDimChange
      TagDispatch = 0
      DataType = 'GCCATEGORIEDIM'
    end
    object cbControleDoublon: TCheckBox
      Left = 388
      Top = 10
      Width = 177
      Height = 17
      Caption = 'Contr'#244'le des doublons'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous sauvegarder vos modifications ?;Q;YNC;Y;' +
        'C'
      
        '1;?caption?;Il n'#39'existe pas de grille dans cette dimension !;W;O' +
        ';O;O'
      '2;?caption?;Chaine non trouv'#233'e dans cette dimension !;W;O;O;O'
      '3;?caption?;Chaine non trouv'#233'e !;W;O;O;O'
      '4;?caption?;Sauvegarde des dimensions non effectu'#233'e !;W;O;O;O'
      
        '5;?caption?;Sauvegarde du d'#233'placement des dimensions non effectu' +
        #233'e !;W;O;O;O'
      
        '6;?caption?;Voulez-vous sortir sans sauvegarder vos modification' +
        's ?;Q;YN;Y;C'
      ' ')
    Left = 124
    Top = 164
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 176
    Top = 224
  end
  object PopLigCol: TPopupMenu
    Left = 384
    Top = 132
    object InsereTaille: TMenuItem
      Caption = 'Ins'#233'rer une taille'
    end
    object Supprimetaille: TMenuItem
      Caption = 'Supprimer une taille'
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object TriAsc: TMenuItem
      Caption = 'Trier par ordre croissant'
    end
    object TriDesc: TMenuItem
      Caption = 'Trier par ordre d'#233'croissant'
    end
  end
end

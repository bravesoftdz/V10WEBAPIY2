inherited FEchgImp_Assistant: TFEchgImp_Assistant
  Left = 271
  Top = 185
  VertScrollBar.Range = 0
  BorderStyle = bsSingle
  Caption = 'Importation'
  ClientHeight = 418
  ClientWidth = 493
  OldCreateOrder = True
  OnConstrainedResize = FormConstrainedResize
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 351
    Width = 493
    Height = 67
    inherited PBouton: TToolWindow97
      ClientHeight = 63
      ClientWidth = 493
      ClientAreaHeight = 63
      ClientAreaWidth = 493
      inherited BValider: TToolbarButton97
        Left = 397
        Hint = 'D'#233'marrer l'#39'importation'
        AllowAllUp = False
      end
      inherited BFerme: TToolbarButton97
        Left = 429
        AllowAllUp = False
      end
      inherited HelpBtn: TToolbarButton97
        Left = 461
        AllowAllUp = False
      end
      inherited BDelete: TToolbarButton97
        Left = 108
        AllowAllUp = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 357
      end
      object btnParametres: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Param'#232'tres'
        DisplayMode = dmGlyphOnly
        Caption = 'Param'#232'tres'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = btnParametresClick
        GlobalIndexImage = 'Z0008_S16G1'
      end
      object btnTables: TToolbarButton97
        Left = 37
        Top = 2
        Width = 39
        Height = 27
        Hint = 'Tables de correspondance'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopTables
        Caption = 'Actions compte'
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0611_S16G1'
      end
      object BRapport: TToolbarButton97
        Tag = 21004
        Left = 77
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Afficher le compte rendu'
        DisplayMode = dmGlyphOnly
        Caption = 'Restaurer un script'
        Flat = False
        OnClick = BRapportClick
        GlobalIndexImage = 'Z0507_S16G1'
      end
      object BKit: TToolbarButton97
        Tag = 21004
        Left = 228
        Top = 2
        Width = 29
        Height = 27
        Hint = 'Importer un script'
        DisplayMode = dmGlyphOnly
        Caption = 'Kit'
        Flat = False
        OnClick = BKitClick
        GlobalIndexImage = 'Z1100_S16G1'
      end
      object bValider2: TToolbarButton97
        Left = 397
        Top = 18
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
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
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        Visible = False
        OnClick = BValiderClick
        IsControl = True
      end
    end
  end
  object PANELREGROUPEMENT: THPanel [1]
    Left = 0
    Top = 0
    Width = 493
    Height = 351
    Align = alClient
    BevelOuter = bvNone
    FullRepaint = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HPanel1: THPanel
      Left = 0
      Top = 0
      Width = 493
      Height = 225
      Align = alTop
      BevelOuter = bvNone
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object panAuto: TPanel
        Left = 0
        Top = 0
        Width = 493
        Height = 80
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        DesignSize = (
          493
          80)
        object HLabel3: THLabel
          Left = 14
          Top = 15
          Width = 40
          Height = 13
          Caption = 'Corbeille'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel4: THLabel
          Left = 14
          Top = 47
          Width = 38
          Height = 13
          Caption = 'Masque'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Hcme_Corbeille: THCritMaskEdit
          Left = 88
          Top = 10
          Width = 249
          Height = 21
          Hint = 'Corbeille des fichiers de donn'#233'es '#224' importer'
          Anchors = []
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnChange = cmeNomFichChange
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
        object Hcme_Masque: THCritMaskEdit
          Left = 88
          Top = 42
          Width = 249
          Height = 21
          Hint = 'Masque des fichiers de donn'#233'es '#224' importer'
          Anchors = []
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          Text = 'Auto-12345.txt'
          OnChange = cmeNomFichChange
          TagDispatch = 0
        end
      end
      object Panel3: TPanel
        Left = 0
        Top = 80
        Width = 493
        Height = 72
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 1
        DesignSize = (
          493
          72)
        object HLABEL2: THLabel
          Left = 14
          Top = 8
          Width = 33
          Height = 13
          Caption = 'Origine'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel1: THLabel
          Left = 14
          Top = 31
          Width = 31
          Height = 13
          Caption = 'Fichier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object btnSelectScriptsReco: TToolbarButton97
          Left = 371
          Top = 3
          Width = 22
          Height = 22
          Hint = 'Param'#232'tres de reconnaissance'
          Anchors = []
          Flat = False
          ParentShowHint = False
          ShowHint = True
          Visible = False
          GlobalIndexImage = 'Z0025_S16G1'
        end
        object btnModeOp: TToolbarButton97
          Left = 410
          Top = 3
          Width = 22
          Height = 22
          Hint = 'Visualisation du mode op'#233'ratoire'
          Anchors = []
          Flat = False
          OnClick = btnModeOpClick
          GlobalIndexImage = 'M0097_S16G1'
        end
        object btnVisuFciDonnees: TToolbarButton97
          Left = 410
          Top = 25
          Width = 22
          Height = 22
          Hint = 'Visualisation du fichier de donn'#233'es '#224' importer'
          Anchors = []
          Flat = False
          OnClick = btnVisuFciDonneesClick
          GlobalIndexImage = 'Z0061_S16G1'
        end
        object Hcme_Solutions: THCritMaskEdit
          Left = 88
          Top = 4
          Width = 249
          Height = 21
          Hint = 'Origine du fichier '#224' importer'
          Anchors = []
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ParentShowHint = False
          ReadOnly = True
          ShowHint = True
          TabOrder = 0
          OnChange = Hcme_SolutionsChange
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = Hcme_SolutionsElipsisClick
          ControlerDate = True
        end
        object cmeNomFich: THCritMaskEdit
          Left = 88
          Top = 27
          Width = 249
          Height = 21
          Hint = 'Fichier de donn'#233'es '#224' importer'
          Anchors = []
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnChange = cmeNomFichChange
          TagDispatch = 0
          DataType = 'OPENFILE(*.*;*.TXT;*.DMX)'
          ElipsisButton = True
        end
      end
      object panDateBal: TPanel
        Left = 0
        Top = 152
        Width = 493
        Height = 41
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 2
        DesignSize = (
          493
          41)
        object HLabDate: THLabel
          Left = 14
          Top = 16
          Width = 54
          Height = 13
          Caption = 'Balance au'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object cmeDate: THCritMaskEdit
          Left = 88
          Top = 12
          Width = 249
          Height = 21
          Hint = 'Date d'#39'int'#233'gration'
          Anchors = []
          EditMask = '!99/99/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          Text = '  /  /    '
          TagDispatch = 0
          OpeType = otDate
          ElipsisButton = True
          ElipsisAutoHide = True
          ControlerDate = True
        end
      end
    end
    object hgVariables: THGrid
      Left = 8
      Top = 198
      Width = 441
      Height = 164
      Hint = 'Variables.'
      ColCount = 3
      DefaultColWidth = 195
      DefaultRowHeight = 16
      FixedCols = 0
      RowCount = 10
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goEditing]
      TabOrder = 1
      SortedCol = -1
      Titres.Strings = (
        'Nom'
        'Valeur')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = -1
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        195
        195
        195)
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 192
    Top = 420
  end
  object PopTables: TPopupMenu
    Left = 145
    Top = 416
    object mConvCptes: TMenuItem
      Caption = 'Correspondance des &Comptes'
      Hint = 'Modification de la table de correspondance des comptes'
      OnClick = mTablesConversionClick
    end
    object mConvJnx: TMenuItem
      Caption = 'Correspondance des &Journaux'
      Hint = 'Modification de la table de correspondance des journaux'
      OnClick = mTablesConversionClick
    end
  end
  object bfKit: TBrowseFolder
    Title = 'Ouvrir un kit'
    Flags = [bfFileSysDirsOnly]
    ShowPathInStatusArea = False
    CustomButtonVisible = True
    CustomButtonCaption = 'Kit'
    SyncCustomButton = False
    Left = 100
    Top = 416
  end
end

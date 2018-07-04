inherited FEXPGRP: TFEXPGRP
  Left = 237
  Top = 137
  Width = 787
  Height = 460
  HelpContext = 1400
  Caption = 'G'#233'n'#233'ration multiple'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 391
    Width = 779
    inherited PBouton: TToolWindow97
      ClientWidth = 779
      ClientAreaWidth = 779
      inherited BValider: TToolbarButton97
        Left = 683
        ParentShowHint = False
        ShowHint = True
      end
      inherited BFerme: TToolbarButton97
        Left = 715
        ParentShowHint = False
        ShowHint = True
      end
      inherited HelpBtn: TToolbarButton97
        Tag = 1400
        Left = 747
        ParentShowHint = False
        ShowHint = True
      end
      inherited bDefaire: TToolbarButton97
        Left = 131
        Top = 3
        Height = 26
      end
      inherited Binsert: TToolbarButton97
        Top = 3
      end
      inherited BDelete: TToolbarButton97
        Top = 3
        ParentShowHint = False
        ShowHint = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 651
        ParentShowHint = False
        ShowHint = True
        Visible = True
        OnClick = BImprimerClick
      end
      object ToolbarButton971: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Vider la liste'
        AllowAllUp = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bDefaireClick
        GlobalIndexImage = 'Z0080_S16G1'
        IsControl = True
      end
      object BVariables: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Saisie des variables'
        AllowAllUp = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BVariablesClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bExport: TToolbarButton97
        Left = 100
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Exporter la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Exporter'
        Flat = False
        Layout = blGlyphTop
        OnClick = bExportClick
        GlobalIndexImage = 'Z0724_S16G1'
      end
    end
  end
  object PageControl1: TPageControl [1]
    Left = 0
    Top = 0
    Width = 779
    Height = 391
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = 'Emission'
      object GD: THGrid
        Left = 0
        Top = 86
        Width = 771
        Height = 277
        Align = alClient
        ColCount = 6
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 2
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
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
          'Fin deTraitement '
          'Variables')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        ColCombo = 1
        ValCombo = ComboScript
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
        ColWidths = (
          183
          129
          172
          78
          102
          98)
      end
      object ComboScript: THValComboBox
        Left = 176
        Top = 80
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        Visible = False
        OnChange = ComboScriptChange
        TagDispatch = 0
        DisableTab = True
      end
      object GroupBox1: TGroupBox
        Left = 0
        Top = 0
        Width = 771
        Height = 86
        Align = alTop
        TabOrder = 2
        DesignSize = (
          771
          86)
        object Label25: TLabel
          Left = 6
          Top = 12
          Width = 42
          Height = 13
          Caption = 'Domaine'
        end
        object Label1: TLabel
          Left = 272
          Top = 12
          Width = 90
          Height = 13
          Caption = 'R'#233'pertoire d'#39'entr'#233'e'
        end
        object Label2: TLabel
          Left = 6
          Top = 39
          Width = 102
          Height = 13
          Caption = 'Extension des fichiers'
        end
        object BValide: TToolbarButton97
          Left = 722
          Top = 54
          Width = 24
          Height = 22
          Hint = 'Reconnaissance par Extension'
          Alignment = taRightJustify
          Anchors = [akRight]
          DisplayMode = dmGlyphOnly
          Caption = 'Ouvrir'
          ParentShowHint = False
          ShowHint = True
          OnClick = BValideClick
          GlobalIndexImage = 'Z0217_S16G1'
        end
        object Label3: TLabel
          Left = 272
          Top = 39
          Width = 92
          Height = 13
          Caption = 'R'#233'pertoire de sortie'
        end
        object Label12: TLabel
          Left = 6
          Top = 65
          Width = 78
          Height = 13
          Caption = 'Correspondance'
        end
        object Label35: TLabel
          Left = 272
          Top = 65
          Width = 53
          Height = 13
          Caption = 'Commande'
        end
        object Reconnaissance: TToolbarButton97
          Left = 735
          Top = 8
          Width = 30
          Height = 21
          Hint = 'Reconnaissance'
          Anchors = [akTop, akRight]
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Layout = blGlyphBottom
          Margin = 0
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          OnClick = ReconnaissanceClick
          GlobalIndexImage = 'Z0313_S24G1'
        end
        object SelectScriptsReco: TToolbarButton97
          Left = 715
          Top = 8
          Width = 20
          Height = 21
          Hint = 'Param'#232'tres de reconnaissance'
          Anchors = [akTop, akRight]
          ParentShowHint = False
          ShowHint = True
          OnClick = SelectScriptsRecoClick
          GlobalIndexImage = 'Z0025_S16G1'
        end
        object ChoixRep: THCritMaskEdit
          Left = 378
          Top = 8
          Width = 322
          Height = 21
          TabOrder = 0
          OnExit = ChoixRepExit
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
        object EExtention: TEdit
          Left = 136
          Top = 35
          Width = 123
          Height = 21
          TabOrder = 1
          Text = '*.*'
        end
        object RepSortie: THCritMaskEdit
          Left = 378
          Top = 35
          Width = 322
          Height = 21
          TabOrder = 2
          OnExit = RepSortieExit
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
        object Profile: THCritMaskEdit
          Left = 136
          Top = 61
          Width = 124
          Height = 21
          MaxLength = 35
          TabOrder = 3
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = ProfileElipsisClick
        end
        object EdShellExecute: THCritMaskEdit
          Left = 378
          Top = 61
          Width = 322
          Height = 21
          TabOrder = 4
          TagDispatch = 0
          DataType = 'OPENFILE(*.BAT;*.EXE;*.*)'
          ElipsisButton = True
        end
        object Domaine: THValComboBox
          Left = 136
          Top = 8
          Width = 121
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 5
          TagDispatch = 0
          ComboWidth = 300
        end
      end
    end
  end
end

inherited FNewListeDistrib: TFNewListeDistrib
  Left = 351
  Top = 169
  Width = 410
  Height = 399
  Caption = 'Listes de distribution'
  OldCreateOrder = True
  ShowHint = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 330
    Width = 402
    inherited PBouton: TToolWindow97
      ClientWidth = 402
      ClientAreaWidth = 402
      ParentShowHint = False
      ShowHint = True
      inherited BValider: TToolbarButton97
        Left = 306
        ModalResult = 0
        ParentShowHint = False
        ShowHint = True
      end
      inherited BFerme: TToolbarButton97
        Left = 338
      end
      inherited HelpBtn: TToolbarButton97
        Left = 370
      end
      inherited bDefaire: TToolbarButton97
        Left = 108
      end
      inherited Binsert: TToolbarButton97
        Hint = 'Nouveau Email'
        ParentShowHint = False
        ShowHint = True
        Visible = True
      end
      inherited BDelete: TToolbarButton97
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 274
      end
      object BAjouterDest: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ajouter des destinataires'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bAjouterDestClick
        GlobalIndexImage = 'Z2213_S16G1'
        IsControl = True
      end
    end
  end
  object Panel1: TPanel [1]
    Left = 0
    Top = 0
    Width = 402
    Height = 330
    Align = alClient
    TabOrder = 1
    object GrilleElementListeDistrib: THGrid
      Left = 1
      Top = 81
      Width = 400
      Height = 248
      Align = alClient
      ColCount = 2
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 1
      OnDblClick = ModifierEmail
      SortedCol = -1
      Titres.Strings = (
        'Code ou email'
        'Nom usuel')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        64
        64)
    end
    object GroupBox1: TGroupBox
      Left = 1
      Top = 1
      Width = 400
      Height = 80
      Align = alTop
      TabOrder = 0
      object HLabel1: THLabel
        Left = 8
        Top = 21
        Width = 22
        Height = 13
        Caption = 'Nom'
      end
      object HLabel2: THLabel
        Left = 8
        Top = 47
        Width = 53
        Height = 13
        Caption = 'Description'
      end
      object CheckBoxPrive: TCheckBox
        Left = 336
        Top = 16
        Width = 57
        Height = 17
        Caption = 'Priv'#233
        TabOrder = 1
      end
      object DescriptionListe: TEdit
        Left = 80
        Top = 42
        Width = 297
        Height = 21
        MaxLength = 100
        TabOrder = 2
      end
      object NomListe: THCritMaskEdit
        Left = 80
        Top = 16
        Width = 161
        Height = 21
        MaxLength = 25
        TabOrder = 0
        OnKeyPress = NomListeKeyPress
        TagDispatch = 0
      end
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 172
    Top = 400
  end
end

inherited FRechercherDestinataire: TFRechercherDestinataire
  Left = 113
  Top = 95
  Width = 800
  Height = 598
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Rechercher destinataire'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 527
    Width = 792
    Height = 37
    inherited PBouton: TToolWindow97
      ClientHeight = 33
      ClientWidth = 792
      ClientAreaHeight = 33
      ClientAreaWidth = 792
      inherited BValider: TToolbarButton97
        Left = 696
      end
      inherited BFerme: TToolbarButton97
        Left = 728
      end
      inherited HelpBtn: TToolbarButton97
        Left = 760
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 664
      end
      object BRechercher: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Flat = False
        Layout = blGlyphTop
        OnClick = BRechercherClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object lAide: THLabel
        Left = 44
        Top = 10
        Width = 214
        Height = 13
        Caption = '(ou taper la premi'#232're lettre du nom recherch'#233')'
      end
    end
  end
  object PageRecherche: TPageControl [1]
    Left = 0
    Top = 0
    Width = 544
    Height = 527
    ActivePage = OngletCollaborateur
    Align = alClient
    TabOrder = 1
    OnExit = ModifierCritereRepertoire
    object OngletCollaborateur: TTabSheet
      Caption = 'Collaborateur'
      object GroupBox1: TGroupBox
        Left = 0
        Top = 0
        Width = 536
        Height = 137
        Align = alTop
        TabOrder = 0
        object LCode: THLabel
          Left = 8
          Top = 20
          Width = 25
          Height = 13
          Caption = 'Code'
        end
        object LNomCollaborateur: THLabel
          Left = 8
          Top = 44
          Width = 22
          Height = 13
          Caption = 'Nom'
        end
        object LLogin: THLabel
          Left = 8
          Top = 68
          Width = 26
          Height = 13
          Caption = 'Login'
        end
        object LGroupe: THLabel
          Left = 8
          Top = 92
          Width = 35
          Height = 13
          Caption = 'Groupe'
        end
        object BCritereCollaborateur: TToolbarButton97
          Left = 496
          Top = 81
          Width = 27
          Height = 26
          OnClick = ModifierCritereCollaborateur
          GlobalIndexImage = 'Z0217_S16G1'
        end
        object Code: THCritMaskEdit
          Left = 80
          Top = 14
          Width = 150
          Height = 21
          TabOrder = 0
          TagDispatch = 0
        end
        object NomCollaborateur: THCritMaskEdit
          Left = 80
          Top = 38
          Width = 150
          Height = 21
          TabOrder = 1
          TagDispatch = 0
        end
        object Groupe: THCritMaskEdit
          Left = 80
          Top = 86
          Width = 150
          Height = 21
          TabOrder = 2
          TagDispatch = 0
        end
        object Login: THCritMaskEdit
          Left = 80
          Top = 62
          Width = 150
          Height = 21
          TabOrder = 3
          TagDispatch = 0
        end
      end
      object GrilleCollaborateur: THGrid
        Left = 0
        Top = 137
        Width = 536
        Height = 362
        Align = alClient
        ColCount = 3
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 1
        SortedCol = -1
        Titres.Strings = (
          'Code'
          'Login'
          'Utilisateur')
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          64
          64
          64)
      end
    end
    object OngletAnnuaire: TTabSheet
      Caption = 'Annuaire'
      ImageIndex = 1
      object GroupBox2: TGroupBox
        Left = 0
        Top = 0
        Width = 536
        Height = 137
        Align = alTop
        TabOrder = 0
        object LDossier: THLabel
          Left = 8
          Top = 20
          Width = 35
          Height = 13
          Caption = 'Dossier'
        end
        object LNomAnnuaire: THLabel
          Left = 8
          Top = 44
          Width = 22
          Height = 13
          Caption = 'Nom'
        end
        object LVille: THLabel
          Left = 8
          Top = 68
          Width = 19
          Height = 13
          Caption = 'Ville'
        end
        object LFonction: THLabel
          Left = 8
          Top = 92
          Width = 41
          Height = 13
          Caption = 'Fonction'
        end
        object BCritereAnnuaire: TToolbarButton97
          Left = 496
          Top = 81
          Width = 27
          Height = 26
          OnClick = ModifierCritereAnnuaire
          GlobalIndexImage = 'Z0217_S16G1'
        end
        object Fonction: THCritMaskEdit
          Left = 80
          Top = 86
          Width = 150
          Height = 21
          TabOrder = 3
          TagDispatch = 0
        end
        object Ville: THCritMaskEdit
          Left = 80
          Top = 62
          Width = 150
          Height = 21
          TabOrder = 2
          TagDispatch = 0
        end
        object NomAnnuaire: THCritMaskEdit
          Left = 80
          Top = 38
          Width = 150
          Height = 21
          TabOrder = 1
          TagDispatch = 0
        end
        object Dossier: THCritMaskEdit
          Left = 80
          Top = 14
          Width = 150
          Height = 21
          TabOrder = 0
          OnExit = ModifierCritereAnnuaire
          TagDispatch = 0
        end
      end
      object GrilleAnnuaire: THGrid
        Left = 0
        Top = 137
        Width = 536
        Height = 369
        Align = alClient
        ColCount = 8
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 1
        SortedCol = -1
        Titres.Strings = (
          'Nom'
          'Dossier'
          'CP'
          'Ville'
          'Civ'
          'Interloc'
          'Fonction'
          'E-Mail')
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          64
          64
          64
          64
          64
          64
          64
          64)
      end
    end
    object OngletRepertoire: TTabSheet
      Caption = 'R'#233'pertoire personnel'
      ImageIndex = 2
      object GroupBox3: TGroupBox
        Left = 0
        Top = 0
        Width = 536
        Height = 137
        Align = alTop
        TabOrder = 0
        object LNomRepertoire: THLabel
          Left = 8
          Top = 20
          Width = 22
          Height = 13
          Caption = 'Nom'
        end
        object LPrenom: THLabel
          Left = 8
          Top = 44
          Width = 36
          Height = 13
          Caption = 'Pr'#233'nom'
        end
        object LEmail: THLabel
          Left = 8
          Top = 68
          Width = 25
          Height = 13
          Caption = 'Email'
        end
        object LCommentaire: THLabel
          Left = 8
          Top = 92
          Width = 61
          Height = 13
          Caption = 'Commentaire'
        end
        object BCritereRepertoire: TToolbarButton97
          Left = 496
          Top = 81
          Width = 27
          Height = 26
          OnClick = ModifierCritereRepertoire
          GlobalIndexImage = 'Z0217_S16G1'
        end
        object Commentaire: THCritMaskEdit
          Left = 80
          Top = 86
          Width = 150
          Height = 21
          TabOrder = 3
          TagDispatch = 0
        end
        object Email: THCritMaskEdit
          Left = 80
          Top = 62
          Width = 150
          Height = 21
          TabOrder = 2
          TagDispatch = 0
        end
        object Prenom: THCritMaskEdit
          Left = 80
          Top = 38
          Width = 150
          Height = 21
          TabOrder = 1
          TagDispatch = 0
        end
        object NomRepertoire: THCritMaskEdit
          Left = 80
          Top = 14
          Width = 150
          Height = 21
          TabOrder = 0
          TagDispatch = 0
        end
      end
      object GrilleRepertoire: THGrid
        Left = 0
        Top = 137
        Width = 536
        Height = 369
        Align = alClient
        ColCount = 4
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 1
        SortedCol = -1
        Titres.Strings = (
          'Nom'
          'Pr'#233'nom'
          'Civ'
          'E-Mail')
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          64
          64
          64
          64)
      end
    end
    object OngletListeDistribution: TTabSheet
      Caption = 'Liste distribution'
      ImageIndex = 3
      object GrilleListeDistribution: THGrid
        Left = 0
        Top = 0
        Width = 536
        Height = 506
        Align = alClient
        ColCount = 3
        DefaultRowHeight = 18
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 0
        SortedCol = -1
        Titres.Strings = (
          'Nom'
          'Description'
          'Priv'#233)
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = True
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          64
          64
          64)
      end
    end
  end
  object Panel1: TPanel [2]
    Left = 544
    Top = 0
    Width = 248
    Height = 527
    Align = alRight
    TabOrder = 2
    object ListeDiffusion: TComboBox
      Left = 70
      Top = 5
      Width = 175
      Height = 21
      Enabled = False
      ItemHeight = 13
      TabOrder = 0
      Visible = False
    end
    object BDestA: TButton
      Left = 5
      Top = 120
      Width = 60
      Height = 21
      Caption = 'A ->'
      TabOrder = 1
      OnClick = BDestAClick
    end
    object BDestCc: TButton
      Left = 5
      Top = 387
      Width = 60
      Height = 21
      Caption = 'Cc ->'
      TabOrder = 3
      OnClick = BDestCcClick
    end
    object GrilleCc: THGrid
      Left = 70
      Top = 267
      Width = 175
      Height = 258
      ColCount = 2
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 4
      SortedCol = -1
      Titres.Strings = (
        'Type'
        'Nom')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
      ColWidths = (
        64
        64)
    end
    object GrilleA: THGrid
      Left = 70
      Top = 6
      Width = 175
      Height = 258
      ColCount = 2
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 2
      SortedCol = -1
      Titres.Strings = (
        'Type'
        'Nom')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
      ColWidths = (
        64
        64)
    end
    object BSupprimerDestA: TButton
      Left = 5
      Top = 144
      Width = 60
      Height = 21
      Caption = '<-'
      TabOrder = 5
      OnClick = BSupprimerDestAClick
    end
    object BSupprimerDestCc: TButton
      Left = 5
      Top = 411
      Width = 60
      Height = 21
      Caption = '<-'
      TabOrder = 6
      OnClick = BSupprimerDestCcClick
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 504
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 472
    Top = 167
  end
end

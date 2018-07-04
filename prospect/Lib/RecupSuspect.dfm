inherited FRecupSuspect: TFRecupSuspect
  Left = 274
  Top = 279
  HelpContext = 111000337
  Caption = 'Assistant Importation de Suspects'
  ClientHeight = 402
  ClientWidth = 618
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 382
    Width = 82
    Height = 18
  end
  inherited lAide: THLabel
    Top = 353
    Width = 394
  end
  object TINTRO: THLabel [2]
    Left = 11
    Top = 48
    Width = 312
    Height = 45
    AutoSize = False
    WordWrap = True
  end
  inherited bPrecedent: TToolbarButton97
    Left = 349
    Top = 376
  end
  inherited bSuivant: TToolbarButton97
    Left = 429
    Top = 376
  end
  inherited bFin: TToolbarButton97
    Left = 509
    Top = 376
    Width = 96
    Caption = '&Lancer l'#39'import'
    Enabled = False
  end
  inherited bAnnuler: TToolbarButton97
    Left = 189
    Top = 376
  end
  inherited bAide: TToolbarButton97
    Left = 108
    Top = 376
  end
  object bparametre: TToolbarButton97 [8]
    Left = 269
    Top = 376
    Width = 75
    Height = 23
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = '&Description'
    Flat = False
    OnClick = BDescriptionClick
  end
  inherited Plan: THPanel
    Height = 280
  end
  inherited GroupBox1: THGroupBox
    Top = 364
    Width = 585
  end
  inherited P: THPageControl2
    Left = 179
    Top = 4
    Width = 411
    Height = 345
    ActivePage = Recapitulatif
    object Fichier: TTabSheet
      Caption = 'Fichier'
      object HLabel1: THLabel
        Left = 13
        Top = 42
        Width = 336
        Height = 26
        Caption = 
          'S'#233'lectionner le code du param'#233'trage, le fichier des suspects '#224' i' +
          'mporter et un code pour identifier l'#39'origine de l'#39'import :'
        WordWrap = True
      end
      object TRSS_PARSUSPECT: THLabel
        Left = 13
        Top = 74
        Width = 118
        Height = 13
        Caption = 'Code param'#233'trage import'
        FocusControl = RSS_PARSUSPECT
      end
      object TRSS_FICHIER: THLabel
        Left = 13
        Top = 98
        Width = 132
        Height = 13
        Caption = 'Localisation fichier suspects'
        FocusControl = RSS_FICHIER
      end
      object LRSS_PARSUSPECT: TLabel
        Left = 242
        Top = 70
        Width = 165
        Height = 21
        AutoSize = False
        Caption = 'D'#233'signation'
      end
      object HLabel3: THLabel
        Left = 13
        Top = 123
        Width = 90
        Height = 13
        Caption = 'Code origine fichier'
        FocusControl = RSS_FICHIER
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 403
        Height = 41
        Align = alTop
        Caption = 'Import d'#39'un fichier Suspects'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object GBCreationFiches: TGroupBox
        Left = 13
        Top = 147
        Width = 378
        Height = 170
        Caption = 'Cr'#233'ation des fiches'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object CBCreationSuspect: TCheckBox
          Left = 18
          Top = 55
          Width = 329
          Height = 13
          Hint = 'Efface la table Suspect'
          Caption = 'Efface et remplace les suspects ayant le code origine pr'#233'c'#233'dent'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object CBControleProspect: TCheckBox
          Left = 18
          Top = 75
          Width = 329
          Height = 16
          Caption = 'Ne pas importer si une fiche Tiers existe avec le m'#234'me siret'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnClick = CBControleProspectClick
        end
        object CBPremiereLigne: TCheckBox
          Left = 18
          Top = 32
          Width = 297
          Height = 17
          Caption = 'Ne pas importer la premi'#232're ligne de description du fichier '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object CBMajTiers: TCheckBox
          Left = 32
          Top = 96
          Width = 229
          Height = 17
          Caption = 'Mettre '#224' jour la fiche Tiers '#224' partir de l'#39'import'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnClick = CBMajTiersClick
        end
        object GBTiers: TGroupBox
          Left = 32
          Top = 132
          Width = 321
          Height = 33
          Enabled = False
          TabOrder = 5
          object CbProspect: TCheckBox
            Left = 8
            Top = 10
            Width = 69
            Height = 17
            Hint = 'Mise '#224' jour des fiches prospect'
            Caption = '&Prospect'
            Color = clBtnFace
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 0
          end
          object CbClient: TCheckBox
            Left = 112
            Top = 10
            Width = 65
            Height = 17
            Hint = 'Mise '#224' jour des fiches client'
            Caption = '&Client'
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 1
          end
          object CbAutre: TCheckBox
            Left = 216
            Top = 10
            Width = 57
            Height = 17
            Hint = 'Mise '#224' jour des fiches Autre'
            Caption = '&Autre'
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 2
          end
        end
        object CBtest: TCheckBox
          Left = 18
          Top = 12
          Width = 347
          Height = 17
          Hint = 'Importer en mode test'
          Caption = 'Mode Test (pour mise '#224' jour de la table de correspondances)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = CBtestClick
        end
        object CBContact: TCheckBox
          Left = 32
          Top = 116
          Width = 221
          Height = 17
          Caption = 'Mise '#224' jour des contacts'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 6
        end
      end
      object RSS_PARSUSPECT: THCritMaskEdit
        Left = 157
        Top = 70
        Width = 63
        Height = 21
        Enabled = False
        TabOrder = 2
        OnExit = RSS_PARSUSPECTExit
        TagDispatch = 0
      end
      object RSS_FICHIER: THCritMaskEdit
        Left = 157
        Top = 95
        Width = 242
        Height = 21
        Hint = 'Nom du Fichier R'#233'cup'#233'r'#233
        TabOrder = 3
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = RSS_FICHIERElipsisClick
      end
      object OrigineFichier: THValComboBox
        Left = 157
        Top = 119
        Width = 124
        Height = 21
        ItemHeight = 13
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'GCORIGINETIERS'
        DataTypeParametrable = True
      end
    end
    object Apercu: TTabSheet
      Caption = 'Apercu'
      ImageIndex = 1
      object GChamp: THGrid
        Tag = 1
        Left = 0
        Top = 0
        Width = 403
        Height = 317
        Align = alClient
        ColCount = 6
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 100
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
        TabOrder = 0
        SortedCol = 0
        Titres.Strings = (
          'Table suspects'
          'Position'
          'Long.'
          'Ligne 1'
          'Ligne 2'
          'Ligne 3')
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          103
          52
          50
          75
          51
          92)
        RowHeights = (
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18)
      end
    end
    object Criteres: TTabSheet
      Caption = 'Crit'#232'res'
      ImageIndex = 2
      object TNomChamp: THLabel
        Left = 1
        Top = 81
        Width = 33
        Height = 13
        Caption = 'Champ'
      end
      object TLongueur: THLabel
        Left = 1
        Top = 113
        Width = 45
        Height = 13
        Caption = 'Longueur'
      end
      object TOffset: THLabel
        Left = 1
        Top = 144
        Width = 37
        Height = 13
        Caption = 'Position'
      end
      object TCritere: THLabel
        Left = 3
        Top = 55
        Width = 44
        Height = 13
        Caption = 'Crit'#232'res'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TBorneSup: THLabel
        Left = 1
        Top = 173
        Width = 54
        Height = 30
        AutoSize = False
        Caption = 'Borne inf'#233'rieure'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        WordWrap = True
      end
      object TBorneInf: THLabel
        Left = 1
        Top = 208
        Width = 56
        Height = 30
        AutoSize = False
        Caption = 'Borne sup'#233'rieure'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        WordWrap = True
      end
      object TTextCritere: THLabel
        Left = 5
        Top = 3
        Width = 379
        Height = 39
        Caption = 
          'Pour ne pas importer tous les suspects du fichier, vous pouvez l' +
          'es s'#233'lectionner en fonction de trois crit'#232'res. Ces crit'#232'res peuv' +
          'ent '#234'tre un champ pr'#233'sent dans le descriptif ou un champ <Autre>' +
          ' avec indication de la longueur et de la position.'
        WordWrap = True
      end
      object GBPremier: TGroupBox
        Left = 56
        Top = 55
        Width = 97
        Height = 185
        Caption = 'Premier'
        TabOrder = 0
        object HChamp1: THValComboBox
          Left = 4
          Top = 21
          Width = 85
          Height = 21
          ItemHeight = 0
          TabOrder = 0
          OnChange = HChamp1Change
          TagDispatch = 0
        end
        object HLongueur1: THCritMaskEdit
          Left = 4
          Top = 53
          Width = 85
          Height = 21
          TabOrder = 1
          TagDispatch = 0
        end
        object HOffset1: THCritMaskEdit
          Left = 4
          Top = 84
          Width = 85
          Height = 21
          TabOrder = 2
          TagDispatch = 0
        end
        object HBorneInf1: THCritMaskEdit
          Left = 4
          Top = 120
          Width = 85
          Height = 21
          TabOrder = 3
          TagDispatch = 0
        end
        object HBorneSup1: THCritMaskEdit
          Left = 4
          Top = 154
          Width = 85
          Height = 21
          TabOrder = 4
          TagDispatch = 0
        end
      end
      object GBDeuxieme: TGroupBox
        Left = 148
        Top = 55
        Width = 97
        Height = 185
        Caption = 'Deuxi'#232'me'
        TabOrder = 1
        object HChamp2: THValComboBox
          Left = 4
          Top = 20
          Width = 85
          Height = 21
          ItemHeight = 0
          TabOrder = 0
          OnChange = HChamp2Change
          TagDispatch = 0
        end
        object HLongueur2: THCritMaskEdit
          Left = 4
          Top = 52
          Width = 85
          Height = 21
          TabOrder = 1
          TagDispatch = 0
        end
        object HOffset2: THCritMaskEdit
          Left = 4
          Top = 83
          Width = 85
          Height = 21
          TabOrder = 2
          TagDispatch = 0
        end
        object HBorneInf2: THCritMaskEdit
          Left = 4
          Top = 120
          Width = 85
          Height = 21
          TabOrder = 3
          TagDispatch = 0
        end
        object HBorneSup2: THCritMaskEdit
          Left = 4
          Top = 153
          Width = 85
          Height = 21
          TabOrder = 4
          TagDispatch = 0
        end
      end
      object GBTroisieme: TGroupBox
        Left = 244
        Top = 55
        Width = 93
        Height = 185
        Caption = 'Troisi'#232'me'
        TabOrder = 2
        object HChamp3: THValComboBox
          Left = 4
          Top = 20
          Width = 85
          Height = 21
          ItemHeight = 0
          TabOrder = 0
          OnChange = HChamp3Change
          TagDispatch = 0
        end
        object HLongueur3: THCritMaskEdit
          Left = 4
          Top = 52
          Width = 85
          Height = 21
          TabOrder = 1
          TagDispatch = 0
        end
        object HOffset3: THCritMaskEdit
          Left = 4
          Top = 83
          Width = 85
          Height = 21
          TabOrder = 2
          TagDispatch = 0
        end
        object HBorneInf3: THCritMaskEdit
          Left = 4
          Top = 120
          Width = 85
          Height = 21
          TabOrder = 3
          TagDispatch = 0
        end
        object HBorneSup3: THCritMaskEdit
          Left = 4
          Top = 152
          Width = 85
          Height = 21
          TabOrder = 4
          TagDispatch = 0
        end
      end
    end
    object Recapitulatif: TTabSheet
      Caption = 'R'#233'capitulatif'
      ImageIndex = 3
      object PanelFin: TPanel
        Left = 5
        Top = 4
        Width = 363
        Height = 90
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 11
          Top = 8
          Width = 334
          Height = 26
          Caption = 
            'Le param'#233'trage est maintenant correctement renseign'#233' pour permet' +
            'tre le lancement de l'#39'import des suspects.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 11
          Top = 43
          Width = 324
          Height = 39
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent, sinon le bouton Lancer l'#39'import permet de d' +
            #233'buter le traitement.'
          WordWrap = True
        end
      end
      object ListRecap: TListBox
        Left = 5
        Top = 100
        Width = 363
        Height = 160
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
    end
  end
  inherited PanelImage: THPanel
    Left = 5
    Top = 74
    inherited Image: TToolbarButton97
      Left = 8
      Top = 10
    end
  end
  inherited cControls: THListBox
    Left = 88
    Top = 268
    Width = 61
    Height = 17
  end
  inherited Msg: THMsgBox
    Police.Height = -9
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;?caption?;La borne inf'#233'rieure 1 < la borne sup'#233'rieure 1;W;O;O;' +
        'O;'
      
        '3;?caption?;La borne inf'#233'rieure 2 < la borne sup'#233'rieure 2;W;O;O;' +
        'O;'
      '4;?caption?;La longueur ou l'#39'offset 2 doit etre saisie;W;O;O;O;'
      
        '5;?caption?;La borne inf'#233'rieure 3 < la borne sup'#233'rieure 3;W;O;O;' +
        'O;'
      '6;?caption?;La longueur ou l'#39'offset 3 doit etre saisie;W;O;O;O;'
      '7;?caption?;Le fichier '#224' r'#233'cup'#233'rer n'#39'existe pas;W;O;O;O;'
      '8;?caption?;La longueur ou l'#39'offset 1 doit etre saisie;W;O;O;O;'
      '9;?caption?;Traitement de r'#233'cup'#233'ration termin'#233';E;O;O;O;'
      
        '10;?caption?;Le code suspect existe,voulez-vous continuer ?;Q;YN' +
        ';Y;N;'
      '11;?caption?;Fichier inaccessible;W;O;O;O;')
    Left = 59
    Top = 12
  end
  object HRecap: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Oui'
      'Non')
    Left = 88
    Top = 44
  end
  object OpenDialogButton: TOpenDialog
    Left = 112
    Top = 8
  end
end

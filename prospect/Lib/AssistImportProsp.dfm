inherited FAsRecupProsp: TFAsRecupProsp
  Left = 356
  Top = 117
  Caption = 'Préparation Fichier Prospect'
  ClientHeight = 382
  ClientWidth = 612
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 344
  end
  inherited bPrecedent: TToolbarButton97
    Left = 196
    Top = 338
    Width = 65
  end
  inherited bSuivant: TToolbarButton97
    Left = 272
    Top = 338
    Width = 65
  end
  inherited bFin: TToolbarButton97
    Left = 348
    Top = 338
    Width = 65
    Caption = '&Traitement'
  end
  inherited bAnnuler: TToolbarButton97
    Left = 121
    Top = 338
    Width = 65
  end
  inherited bAide: TToolbarButton97
    Left = 41
    Top = 302
    Width = 65
  end
  inherited GroupBox1: TGroupBox
    Top = 323
  end
  inherited P: TPageControl
    Top = -1
    Width = 409
    Height = 319
    ActivePage = Parametre
    object Fichier: TTabSheet
      Caption = 'Fichier'
      object Label2: TLabel
        Left = 20
        Top = 104
        Width = 288
        Height = 13
        Caption = 'Localisation du fichier d'#39'archivage de Prospect II :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label3: TLabel
        Left = 20
        Top = 169
        Width = 360
        Height = 13
        Caption = 
          'Le fichier d'#39'archivage Prospect II sera traité avant l'#39'import PG' +
          'I.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Lborne: TLabel
        Left = 189
        Top = 238
        Width = 9
        Height = 13
        Caption = 'à '
      end
      object Label6: TLabel
        Left = 20
        Top = 208
        Width = 359
        Height = 13
        Caption = 
          'Bornes sur le code prospect pour extraire une partie du fichier ' +
          'sinon Suivant.'
      end
      object PTitre1: TPanel
        Left = 16
        Top = 16
        Width = 337
        Height = 74
        TabOrder = 0
        object Label1: TLabel
          Left = 20
          Top = 15
          Width = 307
          Height = 16
          Alignment = taCenter
          Caption = 'Bienvenue dans l'#39'assistant de récupération  '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
        object Label7: TLabel
          Left = 56
          Top = 43
          Width = 211
          Height = 16
          Alignment = taCenter
          Caption = 'des données de PROSPECT II'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object FichierSav: THCritMaskEdit
        Left = 48
        Top = 135
        Width = 292
        Height = 21
        TabOrder = 1
        Text = 'C:\Cegid'
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = FichierSavElipsisClick
      end
      object PremierProspect: TEdit
        Left = 48
        Top = 234
        Width = 106
        Height = 21
        TabOrder = 2
      end
      object DernierProspect: TEdit
        Left = 233
        Top = 234
        Width = 107
        Height = 21
        TabOrder = 3
      end
      object ProgressionRecup2: TEnhancedGauge
        Left = 10
        Top = 270
        Width = 377
        Height = 14
        TabOrder = 4
        Visible = False
        ForeColor = clNavy
        BackColor = clScrollBar
        Progress = 0
      end
    end
    object Grille: TTabSheet
      Caption = 'Grille'
      ImageIndex = 1
      object GridPro: THGrid
        Left = 8
        Top = 7
        Width = 377
        Height = 250
        DefaultColWidth = 90
        DefaultRowHeight = 16
        FixedCols = 0
        RowCount = 100
        FixedRows = 0
        Options = [goVertLine, goHorzLine, goRangeSelect]
        TabOrder = 0
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 1
        ValCombo = CB_TABPRO
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
      end
      object EnhancedGauge1: TEnhancedGauge
        Left = 10
        Top = 270
        Width = 377
        Height = 14
        TabOrder = 1
        Visible = False
        ForeColor = clNavy
        BackColor = clScrollBar
        Progress = 0
      end
    end
    object Parametre: TTabSheet
      Caption = 'Paramètres'
      ImageIndex = 2
      object Label5: TLabel
        Left = 16
        Top = 114
        Width = 265
        Height = 13
        Caption = 'Code représentant dans les grilles techniques (ex  : D ...)'
      end
      object LGrilleInterlocuteur: TLabel
        Left = 16
        Top = 157
        Width = 264
        Height = 13
        Caption = 'Code fonction dans les grilles Interlocuteurs    (ex  : A ...)'
      end
      object LTypeFiche: TLabel
        Left = 16
        Top = 200
        Width = 227
        Height = 13
        Caption = 'Liste des types Fiche pour reprise du code client'
        FocusControl = TypeFiche
      end
      object ProgressionRecup: TEnhancedGauge
        Left = 8
        Top = 238
        Width = 377
        Height = 14
        Caption = 'ProgressionRecup'
        TabOrder = 0
        ForeColor = clNavy
        BackColor = clScrollBar
        Progress = 0
      end
      object GrilleRepresentant: TEdit
        Left = 363
        Top = 110
        Width = 19
        Height = 21
        MaxLength = 1
        TabOrder = 1
      end
      object GrilleInterlocuteur: TEdit
        Left = 363
        Top = 153
        Width = 19
        Height = 21
        MaxLength = 1
        TabOrder = 2
      end
      object TypeFiche: TEdit
        Left = 260
        Top = 196
        Width = 125
        Height = 21
        TabOrder = 3
      end
      object HPreClient: THRadioGroup
        Left = 16
        Top = 48
        Width = 366
        Height = 46
        Caption = 'Début Compte Client'
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          '  9 : Préfixe Compte Client '
          '  C : Préfixe Compte  Client ')
        TabOrder = 4
        Abrege = False
        Vide = False
        Values.Strings = (
          '9'
          'C')
      end
    end
  end
  object CB_TABPRO: THValComboBox [11]
    Left = 10
    Top = 8
    Width = 90
    Height = 21
    Color = clYellow
    Enabled = False
    ItemHeight = 13
    TabOrder = 5
    Visible = False
    TagDispatch = 0
    DisableTab = True
  end
  object BExtraire: TButton [12]
    Left = 505
    Top = 338
    Width = 65
    Height = 23
    Caption = '&Extraire'
    TabOrder = 6
    OnClick = BExtraireClick
  end
  object BImportGRC: TButton [14]
    Left = 426
    Top = 337
    Width = 65
    Height = 25
    Caption = '&Import GRC'
    TabOrder = 7
    OnClick = BImportGRCClick
  end
  inherited Msg: THMsgBox
    Left = 63
    Top = 12
  end
  object OpenFicSav: TOpenDialog
    FileName = 'prarchiv.sav'
    InitialDir = 'C:\PGIS5\V1\Recup'
    Left = 112
    Top = 12
  end
end

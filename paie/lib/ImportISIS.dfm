inherited FImpIsis: TFImpIsis
  Left = 221
  Top = 199
  HelpContext = 42910
  BorderStyle = bsDialog
  Caption = 'Assistant de r'#233'cup'#233'ration des fichiers d'#39'import'
  ClientHeight = 338
  ClientWidth = 625
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object HelpBtn: TToolbarButton97 [2]
    Left = 593
    Top = 308
    Width = 28
    Height = 27
    Hint = 'Aide'
    HelpContext = 42910001
    DisplayMode = dmGlyphOnly
    Caption = 'Aide'
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
    Spacing = -1
    OnClick = HelpBtnClick
    IsControl = True
  end
  inherited bPrecedent: TToolbarButton97
    Left = 311
    Height = 27
  end
  inherited bSuivant: TToolbarButton97
    Left = 386
    Height = 27
  end
  inherited bFin: TToolbarButton97
    Left = 472
    Height = 27
    Visible = False
  end
  inherited bAnnuler: TToolbarButton97
    Left = 224
    Height = 27
  end
  inherited bAide: TToolbarButton97
    Height = 27
    Enabled = False
    Visible = False
  end
  inherited Plan: THPanel
    Left = 236
    Top = 17
  end
  inherited GroupBox1: THGroupBox
    Width = 624
  end
  inherited P: THPageControl2
    Width = 397
    ActivePage = PFileName
    object PFileName: TTabSheet
      Caption = 'Premi'#232're '#233'tape'
      object tedtFichier: TLabel
        Left = 8
        Top = 192
        Width = 37
        Height = 13
        Caption = '&Fichier  '
        FocusControl = edtFichier
      end
      object Label1: TLabel
        Left = 28
        Top = 8
        Width = 288
        Height = 26
        Caption = 
          'Premi'#232're '#233'tape : renseignez le nom du fichier de sauvegarde que ' +
          'vous voulez importer'
        WordWrap = True
      end
      object edtFichier: THCritMaskEdit
        Left = 72
        Top = 188
        Width = 253
        Height = 21
        Hint = 'Nom du fichier '#224' importer'
        TabOrder = 0
        OnChange = edtFichierChange
        TagDispatch = 0
        DataType = 'OPENFILE(*PGI*.SAV)'
        ElipsisButton = True
      end
      object FTypeFichier: TRadioGroup
        Left = 32
        Top = 52
        Width = 321
        Height = 105
        Caption = 'Type de fichier'
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Fichier CEGID ISIS II'
          'Fichier TDS'
          'Fichier de mouvement TRA'
          'Fichier DADS-U')
        TabOrder = 1
        OnClick = FTypeFichierClick
      end
    end
    object PCodageSal: TTabSheet
      Caption = 'Codage des salari'#233's'
      ImageIndex = 1
      object TLCodage: TLabel
        Left = 28
        Top = 32
        Width = 340
        Height = 26
        Caption = 
          'Pr'#233'cisez la fa'#231'on de coder les salari'#233's (ATTENTION en cas de cod' +
          'age chronologique, vous ne pourrez pas importer deux fois le m'#234'm' +
          'e fichier)'
        WordWrap = True
      end
      object FCodeSal: TRadioGroup
        Left = 48
        Top = 108
        Width = 333
        Height = 101
        ItemIndex = 0
        Items.Strings = (
          'Sans changement du num'#233'ro de salari'#233
          'R'#233'affectation par ordre chronologique'
          'Code Etablissement+Code Salari'#233
          'Code Salari'#233'+Code Etablissement')
        TabOrder = 0
      end
    end
    object PCumuls: TTabSheet
      Caption = 'Correspondance Cumuls'
      ImageIndex = 4
      object Label_cumul: TLabel
        Left = 3
        Top = 109
        Width = 306
        Height = 13
        Caption = 'Cliquez sur <Suivant> pour passer '#224' la r'#233'cup'#233'ration des donn'#233'es'
        WordWrap = True
      end
      object GCumuls: THGrid
        Left = 0
        Top = 32
        Width = 386
        Height = 212
        ColCount = 3
        DefaultColWidth = 30
        DefaultRowHeight = 18
        FixedCols = 2
        RowCount = 20
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing, goAlwaysShowEditor]
        TabOrder = 0
        SortedCol = -1
        Titres.Strings = (
          'Code'
          'Cumul PGI'
          'Correspondance')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        ColCombo = 2
        ValCombo = FCumulDest
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = 13224395
        ColWidths = (
          30
          151
          184)
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
          18)
      end
      object FCumulDest: THValComboBox
        Left = 184
        Top = 52
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        Visible = False
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGTABLETEMPO'
        DisableTab = True
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 389
        Height = 32
        Align = alTop
        BevelOuter = bvNone
        Caption = ' '
        TabOrder = 2
        object BCumuls: TToolbarButton97
          Left = 248
          Top = 2
          Width = 137
          Height = 27
          Hint = 'Nouveau'
          Caption = 'G'#233'rer les cumuls'
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Margin = 2
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BCumulsClick
          GlobalIndexImage = 'Z0797_S16G1'
        end
        object FCumul: TCheckBox
          Left = 24
          Top = 8
          Width = 201
          Height = 17
          Caption = 'R'#233'cup'#233'ration des cumuls'
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
      end
    end
    object PLancer: TTabSheet
      Caption = 'Lancement de la r'#233'cup'#233'ration'
      ImageIndex = 3
      object BLancer: TToolbarButton97
        Left = 72
        Top = 166
        Width = 261
        Height = 27
        Hint = 'Lancer la r'#233'cup'#233'ration'
        Caption = 'lancer la r'#233'cup'#233'ration'
        Flat = False
        OnClick = BLancerClick
        GlobalIndexImage = 'Z0184_S16G1'
      end
      object Label3: TLabel
        Left = 88
        Top = 124
        Width = 230
        Height = 13
        Caption = 'Pour lancer la r'#233'cup'#233'ration, cliquez sur le bouton'
        WordWrap = True
      end
      object ChckBxRem: TCheckBox
        Left = 72
        Top = 44
        Width = 177
        Height = 17
        Caption = 'r'#233'cup'#233'ration des r'#233'mun'#233'rations'
        Enabled = False
        TabOrder = 0
        Visible = False
      end
      object RecupRib: TCheckBox
        Left = 72
        Top = 84
        Width = 285
        Height = 17
        Caption = 'R'#233'cup'#233'ration des RIB uniquement'
        TabOrder = 1
        Visible = False
      end
    end
    object PReport: TTabSheet
      Caption = 'Compte rendu d'#39'importation'
      ImageIndex = 2
      object FReport: TListBox
        Left = 0
        Top = 0
        Width = 386
        Height = 244
        Align = alLeft
        ItemHeight = 13
        TabOrder = 0
      end
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;Attention;Certains cumuls vont '#234'tre perdus, Confirmez ?;Q;YN;N' +
        ';N')
    Left = 19
    Top = 4
  end
  inherited HMTrad: THSystemMenu
    Left = 132
    Top = 3
  end
end

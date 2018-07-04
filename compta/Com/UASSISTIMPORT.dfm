inherited FAssistImport: TFAssistImport
  Left = 280
  Top = 139
  HelpContext = 10001
  Caption = 'Assistant de r'#233'ception ComSx '
  ClientHeight = 501
  ClientWidth = 661
  Menu = MainMenu1
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 1
    Top = 476
    Width = 61
  end
  inherited lAide: THLabel
    Top = 438
    Width = 473
    Height = 12
  end
  inherited bPrecedent: TToolbarButton97
    Left = 425
    Top = 471
  end
  inherited bSuivant: TToolbarButton97
    Left = 504
    Top = 471
  end
  inherited bFin: TToolbarButton97
    Left = 582
    Top = 471
  end
  inherited bAnnuler: TToolbarButton97
    Left = 347
    Top = 471
  end
  inherited bAide: TToolbarButton97
    Left = 270
    Top = 471
    Width = 51
    HelpContext = 10001
  end
  object BFiltre: TToolbarButton97 [7]
    Left = 61
    Top = 470
    Width = 58
    Height = 24
    Hint = 'Menu filtre'
    DropdownArrow = True
    Caption = 'Filtre'
    Flat = False
    Layout = blGlyphRight
  end
  object BINI: TToolbarButton97 [8]
    Left = 321
    Top = 471
    Width = 23
    Height = 22
    Hint = 'G'#233'n'#233'ration du fichier de commande'
    Glyph.Data = {
      36030000424D3603000000000000360000002800000010000000100000000100
      1800000000000003000000000000000000000000000000000000C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0C080808000000000000000000000000000000000000000
      0000000000000000000000000000000000808080C0C0C0C0C0C0000000C0C0C0
      8080808080808080808080808080808080808080808080808080808080808080
      80000000C0C0C0C0C0C0000000FFFFFFC0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0C0C0C0C0C0808080000000C0C0C0C0C0C0000000FFFFFF
      C0C0C0C0C0C0C0C0C000FF0080FF0080FF000000FF0000FF0000FFC0C0C08080
      80000000000000000000000000FFFFFFC0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0C0C0C0C0C0808080000000FFFFFF000000000000FFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC0C0
      C0000000FFFFFF00000080808000000000000000000000000000000000000000
      0000000000000000000000000000000000808080FFFFFF000000C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFF000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0000000FFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFF000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0000000FFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFF000000000000000000000000000000C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFF
      FFFFFFFF000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0000000FFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFF000000FFFFFF000000C0C0C0C0C0C0C0C0C0C0C0C0
      C0C0C0C0C0C0C0C0C0000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000000000
      00C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C000000000000000
      0000000000000000000000000000C0C0C0C0C0C0C0C0C0C0C0C0}
    GlyphMask.Data = {00000000}
    ParentShowHint = False
    ShowHint = True
    OnClick = BINIClick
  end
  inherited Plan: THPanel
    Top = 42
  end
  inherited GroupBox1: THGroupBox
    Top = 451
    Width = 746
  end
  inherited P: THPageControl2
    Left = 164
    Top = 5
    Width = 493
    Height = 431
    ActivePage = Mail
    object Mail: TTabSheet
      Caption = 'Mail'
      object Label2: TLabel
        Left = 8
        Top = 160
        Width = 109
        Height = 13
        Caption = 'Adresse de messagerie'
        FocusControl = FEMail
        Transparent = True
      end
      object Bevel5: TBevel
        Left = 130
        Top = 27
        Width = 214
        Height = 8
        Shape = bsTopLine
      end
      object Label11: TLabel
        Left = 7
        Top = 333
        Width = 411
        Height = 13
        Caption = 
          'ATTENTION : le fichier de rapport sera envoy'#233' sous forme de pi'#232'c' +
          'e jointe au message.'
        FocusControl = FEMail
        Transparent = True
      end
      object Label12: TLabel
        Left = 8
        Top = 181
        Width = 93
        Height = 13
        Caption = 'Corps du message :'
        FocusControl = FEMail
        Transparent = True
      end
      object Label6: TLabel
        Left = 8
        Top = 32
        Width = 104
        Height = 13
        Caption = 'Envoyer le rapport par'
      end
      object Label3: TLabel
        Left = 8
        Top = 64
        Width = 124
        Height = 13
        Caption = 'Nom du fichier d'#39#233'change '
        Transparent = True
      end
      object HLABEL2: THLabel
        Left = 8
        Top = 134
        Width = 115
        Height = 13
        Caption = 'Traitement du code stat.'
        Visible = False
      end
      object bClickVisu: TToolbarButton97
        Left = 413
        Top = 59
        Width = 38
        Height = 22
        Hint = 'Contr'#244'le de coh'#233'rence du fichier'
        DropdownArrow = True
        DropdownMenu = POPZ2
        Glyph.Data = {
          66010000424D6601000000000000760000002800000014000000140000000100
          040000000000F0000000C40E0000C40E00001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7777777700007777777777777777777700007777777777777777777700007777
          7777777777777777000077777777777777777777000077000077777000077777
          000070E77707770E77707777000070EEE707770EE7707777000070EEE707770E
          EE707777000070EEE707770EE7E0777700007700007000700000077700007700
          0777777777770777000077770777777777770077000077777007777777777700
          0000777777007707777777700000777777770007777777770000777777777777
          7777777700007777777777777777777700007777777777777777777700007777
          77777777777777770000}
        GlyphMask.Data = {00000000}
      end
      object HLabel1: THLabel
        Left = 128
        Top = 8
        Width = 216
        Height = 13
        Caption = 'Bienvenue dans l'#39'assistant de ComSx '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TListeFichier: THLabel
        Left = 8
        Top = 88
        Width = 116
        Height = 13
        Caption = 'Liste des fichiers '#224' traiter'
        Enabled = False
      end
      object FEMail: THCritMaskEdit
        Left = 139
        Top = 156
        Width = 267
        Height = 21
        TabOrder = 0
        Text = ' '
        OnExit = FEMailExit
        TagDispatch = 0
        ControlerDate = True
      end
      object FHigh: TCheckBox
        Left = 100
        Top = 182
        Width = 181
        Height = 12
        Alignment = taLeftJustify
        Caption = 'Message avec importance haute'
        TabOrder = 1
      end
      object FFile: TRadioButton
        Left = 138
        Top = 30
        Width = 54
        Height = 17
        Caption = 'Fichier'
        Checked = True
        TabOrder = 2
        TabStop = True
        OnClick = FMailClick
      end
      object FMail: TRadioButton
        Left = 197
        Top = 30
        Width = 75
        Height = 17
        Caption = 'Messagerie'
        TabOrder = 3
        OnClick = FMailClick
      end
      object FICHENAME: THCritMaskEdit
        Left = 139
        Top = 60
        Width = 268
        Height = 21
        TabOrder = 4
        OnChange = FICHENAMEChange
        TagDispatch = 0
        DataType = 'OPENFILE(*.TRA;*.ZIP;*.TRT;*.TXT;*.*)'
        ElipsisButton = True
      end
      object CP_STAT: THValComboBox
        Left = 139
        Top = 130
        Width = 195
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        Visible = False
        Items.Strings = (
          'Pas de r'#233'cup'#233'ration'
          'R'#233'cup'#233'r'#233' en table libre mouvement'
          'R'#233'cup'#233'r'#233' en analytique')
        TagDispatch = 0
        Values.Strings = (
          'PAS'
          'TL'
          'ANA')
      end
      object FLISTEFICHIER: THMultiValComboBox
        Left = 139
        Top = 84
        Width = 266
        Height = 21
        Enabled = False
        TabOrder = 6
        Abrege = False
        Complete = False
        OuInclusif = False
      end
      object BNetExpert: TCheckBox
        Left = 356
        Top = 108
        Width = 49
        Height = 17
        Alignment = taLeftJustify
        Caption = 'ASP'
        TabOrder = 7
        Visible = False
        OnClick = BNetExpertClick
      end
      object OPTIONAVANCE: TCheckBox
        Left = 8
        Top = 108
        Width = 144
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Options Avanc'#233'es'
        Enabled = False
        TabOrder = 8
        OnClick = OPTIONAVANCEClick
      end
      object SuppComptable: TCheckBox
        Left = 178
        Top = 108
        Width = 140
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Suppression comptable'
        Enabled = False
        TabOrder = 9
      end
      object BEnvoiMail: TCheckBox
        Left = 284
        Top = 180
        Width = 201
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Envoyer en cas d'#39'anomalie uniquement'
        TabOrder = 10
      end
      object FCorpsMail: THMemo
        Left = 8
        Top = 200
        Width = 473
        Height = 129
        Lines.Strings = (
          'FCorpsMail')
        TabOrder = 11
      end
    end
    object AVANCE: TTabSheet
      Caption = 'Mode Avanc'#233
      ImageIndex = 1
      object HLabel3: THLabel
        Left = 108
        Top = 0
        Width = 104
        Height = 13
        Caption = 'Options Avanc'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel1: TBevel
        Left = 110
        Top = 16
        Width = 97
        Height = 3
        Shape = bsTopLine
      end
      object GroupBox2: TGroupBox
        Left = 47
        Top = 173
        Width = 403
        Height = 65
        Caption = 'Interdire'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object BInterdireCreatCompte: TCheckBox
          Left = 10
          Top = 14
          Width = 210
          Height = 17
          Alignment = taLeftJustify
          Caption = 'La cr'#233'ation en l'#39'absence du param'#233'trage'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = BInterdireCreatCompteClick
        end
        object INTERDIRECREAT: THMultiValComboBox
          Left = 225
          Top = 12
          Width = 170
          Height = 21
          Enabled = False
          HideSelection = False
          TabOrder = 1
          Abrege = False
          Values.Strings = (
            'IGE'
            'IAU'
            'IJA'
            'IET'
            'ISE')
          Items.Strings = (
            'G'#233'n'#233'raux'
            'Auxiliaires'
            'Journaux'
            'Etablissements'
            'Sections')
          Aucun = True
          Complete = False
          OuInclusif = False
        end
        object BGestiondoublon: TCheckBox
          Left = 10
          Top = 40
          Width = 83
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Les doublons'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object BIntegre: TCheckBox
          Left = 99
          Top = 40
          Width = 291
          Height = 17
          Alignment = taLeftJustify
          Caption = ' Int'#233'gration partielle (uniquement les '#233'critures coh'#233'rentes) '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
      end
      object GroupBox3: TGroupBox
        Left = 47
        Top = 130
        Width = 403
        Height = 41
        Caption = 'Conserver'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object BRupturePiece: TCheckBox
          Left = 11
          Top = 16
          Width = 154
          Height = 17
          Hint = 'Respecte la rupture des pi'#232'ces du fichier'
          Alignment = taLeftJustify
          Caption = 'Les ruptures des pi'#232'ces'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object BCalculPiece: TCheckBox
          Left = 204
          Top = 16
          Width = 185
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Les num'#233'ros de pi'#232'ce'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = BCalculPieceClick
        end
      end
      object GroupBox4: TGroupBox
        Left = 47
        Top = 88
        Width = 403
        Height = 43
        Caption = 'Contr'#244'ler'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object BControlParam: TCheckBox
          Left = 10
          Top = 15
          Width = 155
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Les param'#232'tres soci'#233't'#233's'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object BControlEtab: TCheckBox
          Left = 206
          Top = 15
          Width = 182
          Height = 17
          Alignment = taLeftJustify
          Caption = 'L'#39'int'#233'grit'#233' des '#233'tablissements'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
      end
      object GroupBox8: TGroupBox
        Left = 47
        Top = 241
        Width = 403
        Height = 168
        Caption = 'Traitement des '#233'critures'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object Label1: TLabel
          Left = 10
          Top = 69
          Width = 130
          Height = 13
          Caption = 'Tables de correspondance '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BCorresp: TToolbarButton97
          Left = 359
          Top = 62
          Width = 35
          Height = 27
          Hint = 'Tables de correspondances'
          HelpContext = 1710
          DropdownArrow = True
          DropdownMenu = POPZ
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            0400000000000001000000000000000000001000000010000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
            000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
            00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
            F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
            0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
            FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
            FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
            0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
            00333377737FFFFF773333303300000003333337337777777333}
          GlyphMask.Data = {00000000}
          Margin = 3
          NumGlyphs = 2
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          IsControl = True
        end
        object HTYPEIMPORT: THLabel
          Left = 10
          Top = 96
          Width = 135
          Height = 13
          Caption = 'Importer les '#233'critures en type'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 10
          Top = 123
          Width = 215
          Height = 13
          Caption = 'D'#233'coupage journal libre ou bordereau tout les'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label4: TLabel
          Left = 303
          Top = 123
          Width = 40
          Height = 13
          Caption = #233'critures'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BRejet: TCheckBox
          Left = 10
          Top = 19
          Width = 156
          Height = 17
          Alignment = taLeftJustify
          Caption = 'G'#233'n'#233'rer un fichier de rejet'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object BValider: TCheckBox
          Left = 204
          Top = 19
          Width = 184
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Valider les '#233'critures '#224' l'#39'int'#233'gration'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object BCtrP: TCheckBox
          Left = 10
          Top = 43
          Width = 156
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Recalcul des contreparties'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object BCTRLSOCIETE: TCheckBox
          Left = 204
          Top = 43
          Width = 184
          Height = 17
          Alignment = taLeftJustify
          Caption = 'V'#233'rifier le code soci'#233't'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object FCORRESP: THMultiValComboBox
          Left = 152
          Top = 65
          Width = 205
          Height = 21
          HideSelection = False
          TabOrder = 4
          Abrege = False
          DataType = 'TTETABLISSEMENT'
          Values.Strings = (
            'IGE'
            'IAU'
            'IA1'
            'IA2'
            'IA3'
            'IA4'
            'IA5'
            'IPM'
            'IJA'
            'SIS'
            'IET')
          Items.Strings = (
            'G'#233'n'#233'raux'
            'Auxiliaires'
            'Axe analytique 1'
            'Axe analytique 2'
            'Axe analytique 3'
            'Axe analytique 4'
            'Axe analytique 5'
            'Mode de paiement'
            'Journaux'
            'Param'#233'trage SISCOII'
            'Etablissement')
          Aucun = True
          Complete = False
          OuInclusif = False
        end
        object TQUALPIECE: THValComboBox
          Left = 152
          Top = 92
          Width = 204
          Height = 21
          Style = csDropDownList
          ItemHeight = 0
          TabOrder = 5
          OnChange = TQUALPIECEChange
          TagDispatch = 0
          DataType = 'TTQUALPIECE'
          DataTypeParametrable = True
        end
        object DecoupLon: TSpinEdit
          Left = 235
          Top = 118
          Width = 56
          Height = 23
          AutoSize = False
          MaxValue = 255
          MinValue = 0
          TabOrder = 6
          Value = 0
        end
        object BDESEQUILIBRE: TCheckBox
          Left = 10
          Top = 145
          Width = 223
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Rejet en cas de d'#233's'#233'quilibre des montants'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 7
        end
      end
      object GroupBox9: TGroupBox
        Left = 47
        Top = 17
        Width = 403
        Height = 72
        Caption = 'Traitement des comptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object Label5: TLabel
          Left = 191
          Top = 45
          Width = 21
          Height = 13
          Caption = 'pour'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object BMAJCPTE: TCheckBox
          Left = 10
          Top = 16
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Mise '#224' jour des comptes'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = BMAJCPTEClick
        end
        object INTERDIRECREATTIERS: TCheckBox
          Left = 207
          Top = 16
          Width = 181
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Interdire la cr'#233'ation des tiers'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object BBlanc: TCheckBox
          Left = 8
          Top = 43
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Prise en compte des zones vides'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = BBlancClick
        end
        object MODIFTIERS: THMultiValComboBox
          Left = 224
          Top = 41
          Width = 170
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          HideSelection = False
          ParentFont = False
          TabOrder = 3
          Abrege = False
          Values.Strings = (
            'T_EAN'
            'T_TABLE0'
            'T_TABLE1'
            'T_TABLE2'
            'T_TABLE3'
            'T_TABLE4'
            'T_TABLE5'
            'T_TABLE6'
            'T_TABLE7'
            'T_TABLE8'
            'T_TABLE9'
            'T_ADRESSE1'
            'T_ADRESSE2'
            'T_ADRESSE3'
            'T_CODEPOSTAL'
            'T_VILLE'
            'T_PAYS'
            'T_ABREGE'
            'T_LANGUE'
            'T_TELEPHONE'
            'T_FAX'
            'T_COMMENTAIRE'
            'T_NIF'
            'T_SIRET'
            'T_APE'
            'T_PRENOM'
            'T_FORMEJURIDIQUE'
            'T_TVAENCAISSEMENT'
            'T_RELANCEREGLEMENT'
            'T_RELANCETRAITE'
            'YTC_TABLELIBRETIERS1'
            'YTC_TABLELIBRETIERS2'
            'YTC_TABLELIBRETIERS3'
            'YTC_TABLELIBRETIERS4'
            'YTC_TABLELIBRETIERS5'
            'YTC_TABLELIBRETIERS6'
            'YTC_TABLELIBRETIERS7'
            'YTC_TABLELIBRETIERS8'
            'YTC_TABLELIBRETIERS9'
            'YTC_TABLELIBRETIERSA'
            'YTC_RESSOURCE1'
            'YTC_RESSOURCE2'
            'YTC_RESSOURCE3')
          Items.Strings = (
            'Code EAN'
            'Table Libre N'#176' 1'
            'Table Libre N'#176' 2'
            'Table Libre N'#176' 3'
            'Table Libre N'#176' 4'
            'Table Libre N'#176' 5'
            'Table Libre N'#176' 6'
            'Table Libre N'#176' 7'
            'Table Libre N'#176' 8'
            'Table Libre N'#176' 9'
            'Table Libre N'#176' 10'
            'Adresse 1'
            'Adresse 2'
            'Adresse 3'
            'Code postal'
            'Ville'
            'Pays'
            'Libell'#233' abr'#233'g'#233
            'Langue'
            'T'#233'l'#233'phone n'#176'1'
            'Fax'
            'Commentaire'
            'Code NIF'
            'Code SIRET'
            'Code NAF'
            'Pr'#233'nom'
            'Forme juridique'
            'Remboursement avoir'
            'Mode relance r'#232'glt.'
            'Mode relance traite'
            'Table Libre compl'#233'mentaire N'#176' 1'
            'Table Libre compl'#233'mentaire N'#176' 2'
            'Table Libre compl'#233'mentaire N'#176' 3'
            'Table Libre compl'#233'mentaire N'#176' 4'
            'Table Libre compl'#233'mentaire N'#176' 5'
            'Table Libre compl'#233'mentaire N'#176' 6'
            'Table Libre compl'#233'mentaire N'#176' 7'
            'Table Libre compl'#233'mentaire N'#176' 8'
            'Table Libre compl'#233'mentaire N'#176' 9'
            'Table Libre compl'#233'mentaire N'#176' 10'
            'Ressource 1'
            'Ressource 2'
            'Ressource 3'
            '')
          Aucun = True
          Complete = False
          OuInclusif = False
        end
      end
    end
    object AVANCE2: TTabSheet
      Caption = 'Mode Avanc'#233' suite'
      ImageIndex = 4
      object HLabel4: THLabel
        Left = 108
        Top = 8
        Width = 104
        Height = 13
        Caption = 'Options Avanc'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel2: TBevel
        Left = 110
        Top = 27
        Width = 97
        Height = 8
        Shape = bsTopLine
      end
      object GComptes: TGroupBox
        Left = 5
        Top = 31
        Width = 471
        Height = 80
        Caption = ' Comptes inexistants, rempla'#231'ement :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TCPTEDEBUT: THLabel
          Left = 8
          Top = 21
          Width = 37
          Height = 13
          Caption = 'G'#233'n'#233'ral'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TCPTEFIN: THLabel
          Left = 8
          Top = 53
          Width = 26
          Height = 13
          Caption = 'Client'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel5: THLabel
          Left = 162
          Top = 21
          Width = 54
          Height = 13
          Caption = 'Fournisseur'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel6: THLabel
          Left = 162
          Top = 53
          Width = 32
          Height = 13
          Caption = 'Salari'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel7: THLabel
          Left = 318
          Top = 21
          Width = 28
          Height = 13
          Caption = 'divers'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CPTEGENE: THCritMaskEdit
          Left = 56
          Top = 17
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZGENERAL'
          ElipsisButton = True
        end
        object CPTEFOUR: THCritMaskEdit
          Left = 219
          Top = 17
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZTFOURN'
          ElipsisButton = True
        end
        object CPTEDIVERS: THCritMaskEdit
          Left = 360
          Top = 17
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZTDIVERS'
          ElipsisButton = True
        end
        object CPTECLIENT: THCritMaskEdit
          Left = 56
          Top = 49
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZTCLIENT'
          ElipsisButton = True
        end
        object CPTESALAIRE: THCritMaskEdit
          Left = 219
          Top = 49
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZTSALARIE'
          ElipsisButton = True
        end
      end
      object GroupBox5: TGroupBox
        Left = 5
        Top = 111
        Width = 470
        Height = 80
        Caption = 'Sections inexistantes, rempla'#231'ement :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object HLabel8: THLabel
          Left = 8
          Top = 21
          Width = 39
          Height = 13
          Caption = 'Axe N'#176'1'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel9: THLabel
          Left = 8
          Top = 53
          Width = 39
          Height = 13
          Caption = 'Axe N'#176'2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel10: THLabel
          Left = 164
          Top = 21
          Width = 39
          Height = 13
          Caption = 'Axe N'#176'3'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel11: THLabel
          Left = 164
          Top = 53
          Width = 39
          Height = 13
          Caption = 'Axe N'#176'4'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TRAxe5: THLabel
          Left = 317
          Top = 21
          Width = 39
          Height = 13
          Caption = 'Axe N'#176'5'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object RAXE1: THCritMaskEdit
          Left = 61
          Top = 17
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZSECTION'
          ElipsisButton = True
        end
        object RAxe2: THCritMaskEdit
          Left = 61
          Top = 49
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZSECTION2'
          ElipsisButton = True
        end
        object RAxe3: THCritMaskEdit
          Left = 219
          Top = 17
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZSECTION3'
          ElipsisButton = True
        end
        object RAxe4: THCritMaskEdit
          Left = 219
          Top = 49
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZSECTION4'
          ElipsisButton = True
        end
        object RAxe5: THCritMaskEdit
          Left = 361
          Top = 17
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZSECTION5'
          ElipsisButton = True
        end
      end
      object GroupBox6: TGroupBox
        Left = 5
        Top = 192
        Width = 269
        Height = 77
        Caption = 'Substitution par le collectif de la fiche '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object HLabel13: THLabel
          Left = 8
          Top = 25
          Width = 118
          Height = 13
          Caption = ' Pour les comptes Clients'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel14: THLabel
          Left = 12
          Top = 51
          Width = 143
          Height = 13
          Caption = 'Pour les comptes Fournisseurs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SUBCOLLCLI: THCritMaskEdit
          Left = 164
          Top = 21
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZGCOLLCLIENT'
          ElipsisButton = True
        end
        object SUBCOLLFOU: THCritMaskEdit
          Left = 164
          Top = 47
          Width = 97
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnChange = FICHENAMEChange
          TagDispatch = 0
          DataType = 'TZGCOLLFOURN'
          ElipsisButton = True
        end
      end
      object GroupBox7: TGroupBox
        Left = 5
        Top = 271
        Width = 271
        Height = 77
        Caption = 'Cr'#233'ation de comptes de tiers... remplacement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object Label8: THLabel
          Left = 14
          Top = 24
          Width = 91
          Height = 13
          Caption = 'Mode de r'#232'glement'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label9: THLabel
          Left = 14
          Top = 50
          Width = 75
          Height = 13
          Caption = 'R'#233'gime de TVA'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object MRDEFAUT: THValComboBox
          Left = 118
          Top = 20
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TTMODEREGLE'
        end
        object REGDEFAUT: THValComboBox
          Left = 118
          Top = 46
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TTREGIMETVA'
        end
      end
      object GroupBox10: TGroupBox
        Left = 277
        Top = 192
        Width = 198
        Height = 156
        Caption = 'Traitement des '#233'critures'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object LBBAP: THLabel
          Left = 4
          Top = 125
          Width = 78
          Height = 13
          Caption = 'G'#233'n'#233'rer les BAP'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TDATEECR1: THLabel
          Left = 9
          Top = 67
          Width = 3
          Height = 13
          FocusControl = DATEECR1
        end
        object LDATEECR1: THLabel
          Left = 4
          Top = 72
          Width = 82
          Height = 13
          Caption = 'G'#233'n'#233'rer '#224' la date'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BEcartChange: TCheckBox
          Left = 4
          Top = 24
          Width = 189
          Height = 17
          Alignment = taLeftJustify
          Caption = 'G'#233'n'#233'ration d'#39#233'cart de change '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = BInterdireCreatCompteClick
        end
        object CBBAP: THValComboBox
          Left = 92
          Top = 121
          Width = 103
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          Vide = True
          VideString = ' '
          DataType = 'CPTYPEVISABAP'
        end
        object BTPayeur: TCheckBox
          Left = 4
          Top = 46
          Width = 189
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Gestion tiers payeurs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = BInterdireCreatCompteClick
        end
        object DATEECR1: THCritMaskEdit
          Left = 92
          Top = 68
          Width = 103
          Height = 21
          EditMask = '!99/99/0000;1;_'
          MaxLength = 10
          TabOrder = 3
          Text = '01/01/1900'
          OnExit = DATEECR1Exit
          TagDispatch = 0
          OpeType = otDate
          ElipsisButton = True
          ElipsisAutoHide = True
          ControlerDate = True
        end
        object CalculdateEche: TCheckBox
          Left = 4
          Top = 96
          Width = 187
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Recalcul des dates d'#39#233'ch'#233'ance'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 4
        end
      end
      object GroupBox11: TGroupBox
        Left = 13
        Top = 347
        Width = 460
        Height = 50
        Caption = 'Mise '#224' jour des soldes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        object BRecalculSolde: TCheckBox
          Left = 305
          Top = 19
          Width = 140
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Mise '#224' jour des soldes '
          Checked = True
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          State = cbChecked
          TabOrder = 0
        end
        object BCalcSoldeLigne: TCheckBox
          Left = 9
          Top = 19
          Width = 139
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Solde '#224' la ligne'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
        end
        object BMajDateDernMvt: TCheckBox
          Left = 161
          Top = 19
          Width = 139
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Date dernier mouvement'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 2
        end
      end
    end
    object Resume: TTabSheet
      Caption = 'Resume'
      ImageIndex = 2
      object Label13: TLabel
        Left = 193
        Top = 12
        Width = 46
        Height = 13
        Caption = 'R'#233'sum'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel6: TBevel
        Left = 195
        Top = 29
        Width = 44
        Height = 8
        Shape = bsTopLine
      end
      object Bevel3: TBevel
        Left = 4
        Top = 56
        Width = 475
        Height = 135
      end
      object FLib1: TLabel
        Left = 32
        Top = 66
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'Traitement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal1: TLabel
        Left = 100
        Top = 66
        Width = 342
        Height = 26
        AutoSize = False
        Caption = 'Import fichier'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        WordWrap = True
      end
      object FLib3: TLabel
        Left = 32
        Top = 90
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'Fichier'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal3: TLabel
        Left = 100
        Top = 90
        Width = 342
        Height = 13
        AutoSize = False
        Caption = 'C:\sfdoedf\dfijhdif.pgi'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object FLib4: TLabel
        Left = 32
        Top = 114
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'Messagerie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal4: TLabel
        Left = 100
        Top = 114
        Width = 342
        Height = 13
        AutoSize = False
        Caption = 'envoi '#224' l'#39'adresse : jdupont@expert-club.fr'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
    object PIMPORT: TTabSheet
      Caption = 'Compte rendu'
      ImageIndex = 3
      object Label14: TLabel
        Left = 146
        Top = 0
        Width = 79
        Height = 13
        Caption = 'Compte rendu'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel7: TBevel
        Left = 146
        Top = 17
        Width = 82
        Height = 8
        Shape = bsTopLine
      end
      object Label15: TLabel
        Left = 190
        Top = 367
        Width = 260
        Height = 17
        AutoSize = False
        Caption = 'Traitement en cours ...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -17
        Font.Name = 'Small Fonts'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LISTEIMPORT: TListBox
        Left = 0
        Top = -5
        Width = 485
        Height = 324
        ItemHeight = 13
        MultiSelect = True
        PopupMenu = POPZ1
        TabOrder = 0
        OnContextPopup = LISTEIMPORTContextPopup
        OnDblClick = LISTEIMPORTDblClick
      end
    end
  end
  object FFiltres: THValComboBox [14]
    Left = 121
    Top = 472
    Width = 147
    Height = 21
    Style = csDropDownList
    Anchors = [akLeft, akTop, akRight]
    ItemHeight = 13
    TabOrder = 5
    TagDispatch = 0
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 56
    Top = 248
    object Generaux1: TMenuItem
      Caption = 'G'#233'n'#233'raux'
      OnClick = Generaux1Click
    end
    object Auxiliaires1: TMenuItem
      Caption = 'Auxiliaires'
      OnClick = Auxiliaires1Click
    end
    object Axeanalytique11: TMenuItem
      Caption = 'Axe analytique1'
      OnClick = Axeanalytique11Click
    end
    object Axeanalytique21: TMenuItem
      Caption = 'Axe analytique 2'
      OnClick = Axeanalytique21Click
    end
    object Axeanalytique31: TMenuItem
      Caption = 'Axe analytique 3'
      OnClick = Axeanalytique31Click
    end
    object Axeanalytique41: TMenuItem
      Caption = 'Axe analytique 4'
      OnClick = Axeanalytique41Click
    end
    object Axeanalytique51: TMenuItem
      Caption = 'Axe analytique 5'
      OnClick = Axeanalytique51Click
    end
    object ModedePaiement1: TMenuItem
      Caption = 'Mode de Paiement'
      OnClick = ModedePaiement1Click
    end
    object Journaux1: TMenuItem
      Caption = 'Journaux'
      OnClick = Journaux1Click
    end
    object ParamtrageSISCOII1: TMenuItem
      Caption = 'Param'#233'trage SISCOII'
      OnClick = ParamtrageSISCOII1Click
    end
    object Etablissement1: TMenuItem
      Caption = 'Etablissement'
      OnClick = Etablissement1Click
    end
  end
  object POPZ1: TPopupMenu
    Left = 17
    Top = 248
    object MenuItem1: TMenuItem
      Caption = 'Zoom sur compte g'#233'n'#233'ral'
      Visible = False
      OnClick = MenuItem1Click
    end
    object Zoomaux: TMenuItem
      Caption = 'Zoom sur compte auxiliaire'
      Visible = False
      OnClick = ZoomauxClick
    end
    object Zoomana: TMenuItem
      Caption = 'Zoom sur la section analytique'
      OnClick = ZoomanaClick
    end
    object Zoomsurlejournal1: TMenuItem
      Caption = 'Zoom sur le journal'
      OnClick = Zoomsurlejournal1Click
    end
    object Zoomsurtablissement1: TMenuItem
      Caption = 'Zoom sur '#233'tablissement'
      OnClick = Zoomsurtablissement1Click
    end
  end
  object MainMenu1: TMainMenu
    Left = 52
    Top = 297
  end
  object POPZ2: TPopupMenu
    AutoPopup = False
    Left = 20
    Top = 323
    object BVOIR: TMenuItem
      Caption = 'Voir le fichier'
      OnClick = bClickVisuClick
    end
    object CTRPIECE: TMenuItem
      Caption = 'Contr'#244'ler les pi'#232'ces'
      OnClick = CTRPIECEClick
    end
  end
end

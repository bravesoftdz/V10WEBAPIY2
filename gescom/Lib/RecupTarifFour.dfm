inherited FRecupTarifFour: TFRecupTarifFour
  Left = 466
  Top = 270
  HelpContext = 110000112
  Caption = 'Assistant r'#233'cup'#233'ration des catalogues fournisseur'
  ClientHeight = 390
  ClientWidth = 592
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 367
  end
  inherited lAide: THLabel
    Top = 337
    Width = 404
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
    Left = 325
    Top = 361
  end
  inherited bSuivant: TToolbarButton97
    Left = 408
    Top = 361
  end
  inherited bFin: TToolbarButton97
    Left = 492
    Top = 361
  end
  inherited bAnnuler: TToolbarButton97
    Left = 242
    Top = 361
  end
  inherited bAide: TToolbarButton97
    Left = 159
    Top = 361
  end
  inherited Plan: THPanel
    Left = 176
    Height = 300
  end
  inherited GroupBox1: THGroupBox
    Top = 349
    Width = 595
  end
  inherited P: THPageControl2
    Left = 176
    Top = 1
    Width = 417
    Height = 328
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object HLabel1: THLabel
        Left = 11
        Top = 44
        Width = 312
        Height = 29
        AutoSize = False
        Caption = 
          'Cet assistant vous guide dans la r'#233'cup'#233'ration de votre catalogue' +
          ' fournisseur'
        WordWrap = True
      end
      object TRGF_PARFOU: THLabel
        Left = 9
        Top = 76
        Width = 102
        Height = 13
        Caption = 'Code du param'#232'trage'
        FocusControl = GRF_PARFOU
      end
      object TGRF_TIERS: THLabel
        Left = 9
        Top = 100
        Width = 82
        Height = 13
        Caption = 'Code Fournisseur'
        FocusControl = GRF_TIERS
      end
      object LGRF_TIERS: THLabel
        Left = 248
        Top = 100
        Width = 85
        Height = 13
        AutoSize = False
        FocusControl = GRF_TIERS
      end
      object TGRF_FICHIER: THLabel
        Left = 9
        Top = 127
        Width = 31
        Height = 13
        Caption = 'Fichier'
        FocusControl = GRF_FICHIER
      end
      object TDATESUP: THLabel
        Left = 9
        Top = 151
        Width = 136
        Height = 13
        Caption = 'Date expiration du catalogue'
        FocusControl = GRF_FICHIER
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 409
        Height = 41
        Align = alTop
        Caption = 'R'#233'cup'#233'ration de catalogues fournisseurs'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -19
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
        Left = 4
        Top = 176
        Width = 405
        Height = 97
        Caption = 'Cr'#233'ation/Mise '#224' jour des fiches'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TCreationCatalogue: THLabel
          Left = 169
          Top = 10
          Width = 232
          Height = 30
          AutoSize = False
          Caption = 'Cr'#233'ation de la fiche catalogue si inexistante ou mise '#224' jour.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object TCreationArticle: THLabel
          Left = 169
          Top = 41
          Width = 232
          Height = 30
          AutoSize = False
          Caption = 'Cr'#233'ation de la fiche article si inexistante ou mise '#224' jour.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object TMAJFamSsFam: THLabel
          Left = 169
          Top = 72
          Width = 232
          Height = 16
          AutoSize = False
          Caption = 'Mise '#224' jour de la famille ou sous-Famille.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object CBCreationCatalogue: TCheckBox
          Left = 10
          Top = 19
          Width = 87
          Height = 13
          Hint = 'Cr'#233'ation fiche Catalogue'
          Caption = 'Catalogues'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object CBCreationArticle: TCheckBox
          Left = 10
          Top = 50
          Width = 79
          Height = 13
          Hint = 'Cr'#233'ation fiche article'
          Caption = 'Articles'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object chkFamSsFam: TCheckBox
          Left = 10
          Top = 72
          Width = 119
          Height = 17
          Caption = 'Famille/Ss-Famille'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
      end
      object GRF_PARFOU: THCritMaskEdit
        Left = 119
        Top = 73
        Width = 126
        Height = 21
        Color = clActiveBorder
        Enabled = False
        TabOrder = 2
        TagDispatch = 0
      end
      object GRF_TIERS: THCritMaskEdit
        Left = 119
        Top = 98
        Width = 126
        Height = 21
        Color = clActiveBorder
        Enabled = False
        TabOrder = 3
        TagDispatch = 0
      end
      object GRF_FICHIER: THCritMaskEdit
        Left = 48
        Top = 123
        Width = 198
        Height = 21
        Hint = 'Nom du Fichier R'#233'cup'#233'r'#233
        TabOrder = 4
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = GRF_FICHIERElipsisClick
      end
      object DATESUP: THCritMaskEdit
        Left = 152
        Top = 147
        Width = 90
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 5
        Text = '  /  /    '
        OnDblClick = DATESUPDblClick
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odFinMois
        ControlerDate = True
      end
      object GRF_MULTIFOU: THCheckbox
        Left = 248
        Top = 96
        Width = 153
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Fournisseurs dans le fichier'
        Enabled = False
        TabOrder = 6
      end
      object CBRecInit: TCheckBox
        Left = 14
        Top = 278
        Width = 125
        Height = 17
        Caption = 'R'#233'cup'#233'ration initiale'
        TabOrder = 7
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object TNomChamp: THLabel
        Left = 1
        Top = 74
        Width = 22
        Height = 13
        Caption = 'Nom'
      end
      object TLongueur: THLabel
        Left = 1
        Top = 106
        Width = 45
        Height = 13
        Caption = 'Longueur'
      end
      object TOffset: THLabel
        Left = 1
        Top = 137
        Width = 45
        Height = 13
        Caption = 'Offset/N'#176
      end
      object TCritere: THLabel
        Left = 3
        Top = 48
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
        Top = 166
        Width = 54
        Height = 30
        AutoSize = False
        Caption = 'Borne Inf'#233'rieure'
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
        Top = 201
        Width = 56
        Height = 30
        AutoSize = False
        Caption = 'Borne Sup'#233'rieure'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        WordWrap = True
      end
      object HLabel2: THLabel
        Left = 11
        Top = 0
        Width = 312
        Height = 25
        AutoSize = False
        Caption = 'S'#233'lection des enregistrements '#224' traiter'
        WordWrap = True
      end
      object GBPremier: TGroupBox
        Left = 56
        Top = 48
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
        Top = 48
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
        Top = 48
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
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object TRecap: THLabel
        Left = 18
        Top = 108
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object PanelFin: TPanel
        Left = 5
        Top = 4
        Width = 329
        Height = 97
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 23
          Top = 8
          Width = 287
          Height = 26
          Caption = 
            'Le param'#232'trage est maintenant correctement renseign'#233' pour permet' +
            'tre le lancement de la  mise '#224' jour des tarifs.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 49
          Width = 285
          Height = 39
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon le bouton Fin permet de d'#233'buter le tra' +
            'itement.'
          WordWrap = True
        end
      end
      object ListRecap: TListBox
        Left = 5
        Top = 126
        Width = 329
        Height = 116
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
    end
  end
  inherited PanelImage: THPanel
    Left = 5
    Top = 74
    Height = 271
    inherited Image: TToolbarButton97
      Left = 8
      Top = 16
      Width = 153
      Height = 153
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      
        '0;?caption?;La borne inf'#233'rieure 1 < la borne sup'#233'rieur 1;W;O;O;O' +
        ';'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;?caption?;La borne inf'#233'rieure 2 < la borne sup'#233'rieur 2;W;O;O;O' +
        ';'
      '3;?caption?;La longueur ou l'#39'offset 2 doit etre saisie;W;O;O;O;'
      
        '4;?caption?;La borne inf'#233'rieure 3 < la borne sup'#233'rieur 3;W;O;O;O' +
        ';'
      '5;?caption?;La longueur ou l'#39'offset 3 doit etre saisie;W;O;O;O;'
      '6;?caption?;Le fichier '#224' r'#233'cup'#233'rer n'#39'existe pas;W;O;O;O;'
      '7;?caption?;La longueur ou l'#39'offset 1 doit etre saisie;W;O;O;O;'
      '8;?caption?;Traitement de r'#233'cup'#233'ration termin'#233';E;O;O;O;')
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
    Left = 96
    Top = 12
  end
  object OpenDialogButton: TOpenDialog
    Left = 124
    Top = 4
  end
end

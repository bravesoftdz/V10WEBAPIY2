inherited FRecopiTarf: TFRecopiTarf
  Left = 239
  Top = 195
  Caption = 'Assistant recopie de tarifs'
  ClientWidth = 605
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 312
  end
  inherited bPrecedent: TToolbarButton97
    Top = 306
  end
  inherited bSuivant: TToolbarButton97
    Top = 306
  end
  inherited bFin: TToolbarButton97
    Top = 306
  end
  inherited bAnnuler: TToolbarButton97
    Top = 306
  end
  inherited bAide: TToolbarButton97
    Top = 306
  end
  inherited Plan: TPanel
    Top = 9
  end
  inherited GroupBox1: TGroupBox
    Top = 290
    Width = 596
  end
  inherited P: TPageControl
    Left = 172
    Width = 421
    Height = 280
    ActivePage = TabSheet2
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object TINTRO: THLabel
        Left = 19
        Top = 48
        Width = 312
        Height = 37
        AutoSize = False
        Caption = 
          'Cet assistant vous guide afin de recopier des tarifs simplement ' +
          'en fonction du type de tarif'
        WordWrap = True
      end
      object Bevel1: TBevel
        Left = 0
        Top = 44
        Width = 397
        Height = 199
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 413
        Height = 45
        Align = alTop
        Caption = 'Recopie des tarifs'
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
      object GBTARIF: TGroupBox
        Left = 20
        Top = 84
        Width = 337
        Height = 69
        Caption = 'Tarif à recopier '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TTYPETARIF: THLabel
          Left = 8
          Top = 20
          Width = 64
          Height = 13
          Caption = '&Type de tarifs'
          FocusControl = TYPETARIF
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TPERIODE: THLabel
          Left = 8
          Top = 42
          Width = 98
          Height = 13
          Caption = '&Période d'#39'application'
          FocusControl = PERIODE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TYPETARIF: THValComboBox
          Left = 169
          Top = 16
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 0
          OnChange = TYPETARIFOnChange
          TagDispatch = 0
          VideString = '<<Aucun>>'
          DataType = 'GCTARIFTYPE1'
        end
        object PERIODE: THValComboBox
          Left = 169
          Top = 38
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 1
          OnChange = PERIODEOnChange
          TagDispatch = 0
          VideString = '<<Aucun>>'
          DataType = 'GCTARIFPERIODE1'
        end
      end
      object GBTYPE: TGroupBox
        Left = 20
        Top = 164
        Width = 337
        Height = 65
        Caption = 'Type de tarif '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object CB_TARFART: TCheckBox
          Left = 8
          Top = 16
          Width = 97
          Height = 17
          Caption = '&Tarif par article'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 0
          OnClick = CheckBoxOnClick
        end
        object CB_TARFCATART: TCheckBox
          Left = 8
          Top = 39
          Width = 141
          Height = 17
          Caption = 'Tarif par catégorie &articles'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 1
          OnClick = CheckBoxOnClick
        end
        object CB_TARFCLI: TCheckBox
          Left = 178
          Top = 16
          Width = 99
          Height = 17
          Caption = 'Tarif par &client'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 2
          OnClick = CheckBoxOnClick
        end
        object CB_TARFCATCLI: TCheckBox
          Left = 178
          Top = 39
          Width = 147
          Height = 17
          Caption = 'Tarif par ca&tégorie clients'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 3
          OnClick = CheckBoxOnClick
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object Bevel3: TBevel
        Left = 0
        Top = 4
        Width = 413
        Height = 246
      end
      object TLISTEET: THLabel
        Left = 7
        Top = 43
        Width = 114
        Height = 13
        Caption = 'Liste des établissements'
      end
      object BFLECHEDROITE2: TToolbarButton97
        Left = 191
        Top = 80
        Width = 28
        Height = 27
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000120B0000120B00001000000000000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333333333333333333333333333333333333333333333333333333333
          3333333333333333333333333333333333333333333FF3333333333333003333
          3333333333773FF3333333333309003333333333337F773FF333333333099900
          33333FFFFF7F33773FF30000000999990033777777733333773F099999999999
          99007FFFFFFF33333F7700000009999900337777777F333F7733333333099900
          33333333337F3F77333333333309003333333333337F77333333333333003333
          3333333333773333333333333333333333333333333333333333333333333333
          3333333333333333333333333333333333333333333333333333}
        NumGlyphs = 2
        OnClick = ClickFlecheDroite2
      end
      object BFLECHEGAUCHE2: TToolbarButton97
        Left = 191
        Top = 108
        Width = 28
        Height = 27
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000120B0000120B00001000000000000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333333333333333333333333333333333333333333333333333333333
          3333333333333FF3333333333333003333333333333F77F33333333333009033
          333333333F7737F333333333009990333333333F773337FFFFFF330099999000
          00003F773333377777770099999999999990773FF33333FFFFF7330099999000
          000033773FF33777777733330099903333333333773FF7F33333333333009033
          33333333337737F3333333333333003333333333333377333333333333333333
          3333333333333333333333333333333333333333333333333333333333333333
          3333333333333333333333333333333333333333333333333333}
        NumGlyphs = 2
        OnClick = ClickFlecheGauche2
      end
      object TRECOPIER: THLabel
        Left = 7
        Top = 16
        Width = 115
        Height = 13
        Caption = 'Etablissement à recopier'
      end
      object TDANS: THLabel
        Left = 225
        Top = 43
        Width = 113
        Height = 13
        Caption = 'Dans les établissements'
      end
      object BFLECHETOUS: TToolbarButton97
        Left = 191
        Top = 141
        Width = 28
        Height = 27
        Glyph.Data = {
          DE000000424DDE0000000000000076000000280000000D0000000D0000000100
          0400000000006800000000000000000000001000000010000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7000777707707777700077770070077770007777090090777000000009909907
          7000099999990990700009999999909900000999999909907000000009909907
          7000777709009077700077770070077770007777077077777000777777777777
          7000}
        OnClick = ClickFlecheTous
      end
      object BFLECHEAUCUN: TToolbarButton97
        Left = 191
        Top = 169
        Width = 28
        Height = 27
        Glyph.Data = {
          DE000000424DDE0000000000000076000000280000000D0000000D0000000100
          04000000000068000000C40E0000C40E00001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7000777770770777700077770070077770007770900907777000770990990000
          0000709909999999000009909999999900007099099999990000770990990000
          0000777090090777700077770070077770007777707707777000777777777777
          7000}
        OnClick = ClickFlecheAucun
      end
      object GLISTE: THGrid
        Left = 7
        Top = 60
        Width = 178
        Height = 158
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        FixedRows = 0
        TabOrder = 0
        OnDblClick = ClickFlecheDroite2
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
      end
      object GCOPIE: THGrid
        Left = 225
        Top = 60
        Width = 178
        Height = 158
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        FixedRows = 0
        TabOrder = 1
        OnDblClick = ClickFlecheGauche2
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
      end
      object MAZE: TCheckBox
        Left = 7
        Top = 227
        Width = 310
        Height = 17
        Caption = 'Suppression au préalable des tarifs dans les établissements.'
        TabOrder = 2
      end
      object ARECOPIER: THValComboBox
        Left = 224
        Top = 12
        Width = 177
        Height = 21
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 413
        Height = 252
        Align = alClient
      end
      object GBPRIX: TGroupBox
        Left = 20
        Top = 12
        Width = 341
        Height = 93
        Caption = 'Variation des prix '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TCOEF: THLabel
          Left = 8
          Top = 24
          Width = 127
          Height = 13
          Caption = '&Application d'#39'un coefficient'
          FocusControl = COEF
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TARRONDIP: THLabel
          Left = 8
          Top = 56
          Width = 85
          Height = 13
          Caption = '&Méthode d'#39'arrondi'
          FocusControl = ARRONDIP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object COEF: THNumEdit
          Left = 180
          Top = 20
          Width = 145
          Height = 21
          Decimals = 2
          Digits = 12
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          ParentFont = False
          ParentShowHint = False
          ShowHint = False
          TabOrder = 0
          UseRounding = True
          Value = 1
          Validate = False
        end
        object ARRONDIP: THValComboBox
          Left = 180
          Top = 52
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          ParentShowHint = False
          ShowHint = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'GCCODEARRONDI'
        end
      end
      object GBREMISE: TGroupBox
        Left = 20
        Top = 120
        Width = 341
        Height = 61
        Caption = 'Variation des pourcentages '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TMONTANT: THLabel
          Left = 8
          Top = 24
          Width = 83
          Height = 13
          Caption = '&Montant à ajouter'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object MONTANT: THNumEdit
          Left = 180
          Top = 20
          Width = 145
          Height = 21
          Decimals = 2
          Digits = 12
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          ParentFont = False
          TabOrder = 0
          UseRounding = True
          Validate = False
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'TabSheet4'
      ImageIndex = 3
      object TRecap: THLabel
        Left = 37
        Top = 104
        Width = 59
        Height = 13
        Caption = 'Récapitulatif'
      end
      object Bevel4: TBevel
        Left = 0
        Top = 0
        Width = 413
        Height = 253
      end
      object PanelFin: TPanel
        Left = 37
        Top = 4
        Width = 329
        Height = 97
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 23
          Top = 8
          Width = 263
          Height = 26
          Caption = 
            'Le paramétrage est maintenant correctement renseigné pour permet' +
            'tre le lancement de la recopie des tarifs.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 49
          Width = 285
          Height = 39
          Caption = 
            'Si vous désirez revoir le paramétrage, il suffit de cliquer sur ' +
            'le bouton Précédent sinon, le bouton Fin permet de débuter le tr' +
            'aitement.'
          WordWrap = True
        end
      end
      object ListRecap: TListBox
        Left = 37
        Top = 120
        Width = 329
        Height = 122
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
      object BFinRcap: TButton
        Left = 376
        Top = 208
        Width = 28
        Height = 27
        Caption = 'Fin'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Visible = False
        OnClick = ClickBFinRecap
      end
    end
  end
  inherited PanelImage: TPanel
    Left = 5
    inherited Image: TToolbarButton97
      Left = 4
    end
  end
  object HMsgErr: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'L'#39'établissement à recopier est obligatoire'
      'Le tarif doit être recopier dans au moins un établissement'
      'Aucun tarif ne correspond à votre séléction')
    Left = 60
    Top = 16
  end
  object HRecap: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Oui'
      'Non'
      'ATTENTION : Tarif(s) non créé(s) !'
      
        'ATTENTION : Ce tarif, en cours de traitement par un autre utilis' +
        'ateur, n'#39'a pas été enregistrée !'
      'enregistrement(s) créé(s).'
      'Etablissement recopié:'
      'Dans les établissements:')
    Left = 132
    Top = 12
  end
end

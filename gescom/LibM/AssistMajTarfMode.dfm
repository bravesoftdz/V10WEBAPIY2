inherited FMajTarfMode: TFMajTarfMode
  Left = 67
  Top = 151
  Caption = 'Mise à jour des tarifs article'
  ClientHeight = 352
  ClientWidth = 730
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 332
  end
  inherited lAide: THLabel
    Top = 276
    Width = 343
  end
  inherited bPrecedent: TToolbarButton97
    Left = 488
    Top = 326
  end
  inherited bSuivant: TToolbarButton97
    Left = 566
    Top = 326
  end
  inherited bFin: TToolbarButton97
    Left = 650
    Top = 326
  end
  inherited bAnnuler: TToolbarButton97
    Left = 401
    Top = 326
  end
  inherited bAide: TToolbarButton97
    Left = 318
    Top = 326
  end
  inherited Plan: TPanel
    Left = 168
    Top = 13
    Width = 561
  end
  inherited GroupBox1: TGroupBox
    Top = 313
    Width = 778
    Height = 8
  end
  inherited P: TPageControl
    Left = 165
    Top = 3
    Width = 563
    Height = 310
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object Bevel1: TBevel
        Left = 0
        Top = 40
        Width = 553
        Height = 241
      end
      object TINTRO: THLabel
        Left = 47
        Top = 44
        Width = 490
        Height = 21
        AutoSize = False
        Caption = 
          'Cet assistant vous guide afin de paramètrer la mise à jour de vo' +
          's tarifs préalablement sélectionnés.'
        WordWrap = True
      end
      object GBDe: TGroupBox
        Left = 4
        Top = 60
        Width = 271
        Height = 213
        Caption = 'A partir de '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object TETABBASE: THLabel
          Left = 8
          Top = 189
          Width = 106
          Height = 13
          Caption = '&Etablissement de base'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object GBPrix: TGroupBox
          Left = 4
          Top = 13
          Width = 263
          Height = 105
          Caption = 'Prix '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = GBPrixOnClick
          OnEnter = GBPrixOnClick
          object RBDPA: TRadioButton
            Left = 4
            Top = 16
            Width = 177
            Height = 17
            Caption = 'Dernier prix d'#39'achat par boutique'
            TabOrder = 0
            OnClick = GBPrixOnClick
          end
          object RBPA: TRadioButton
            Left = 4
            Top = 33
            Width = 113
            Height = 17
            Caption = 'Prix d'#39'achat article'
            TabOrder = 1
            OnClick = GBPrixOnClick
          end
          object RBPVTTC: TRadioButton
            Left = 4
            Top = 68
            Width = 133
            Height = 17
            Caption = 'Prix de vente TTC tarif'
            TabOrder = 2
            OnClick = GBPrixOnClick
          end
          object RBPOURC: TRadioButton
            Left = 188
            Top = 85
            Width = 69
            Height = 17
            Caption = 'Tarif en %'
            TabOrder = 3
            OnClick = GBPrixOnClick
          end
          object RBPVART: TRadioButton
            Left = 4
            Top = 51
            Width = 145
            Height = 17
            Caption = 'Prix de vente TTC article'
            TabOrder = 4
            OnClick = GBPrixOnClick
          end
          object RBPXSAISI: TRadioButton
            Left = 188
            Top = 68
            Width = 65
            Height = 17
            Caption = 'Prix saisi'
            TabOrder = 5
            OnClick = GBPrixOnClick
          end
          object RBAHT: TRadioButton
            Left = 4
            Top = 85
            Width = 133
            Height = 17
            Caption = 'Prix d'#39'achat HT tarif'
            Checked = True
            TabOrder = 6
            TabStop = True
            OnClick = GBPrixOnClick
          end
        end
        object ETABBASE: THValComboBox
          Left = 116
          Top = 185
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
          OnChange = ComboOnChange
          TagDispatch = 0
          Vide = True
          VideString = 'Toutes boutiques'
          DataType = 'TTETABLISSEMENT'
        end
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 553
        Height = 41
        Caption = 'Mise à jour des tarifs'
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
        Left = 8
        Top = 180
        Width = 263
        Height = 61
        Caption = 'Tarifs de base '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TTYPETARIF: THLabel
          Left = 4
          Top = 16
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
          Left = 4
          Top = 38
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
          Left = 113
          Top = 12
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
          OnChange = ComboOnChange
          TagDispatch = 0
          VideString = '<<Aucun>>'
          DataType = 'GCTARIFTYPE1'
        end
        object PERIODE: THValComboBox
          Left = 113
          Top = 34
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
          OnChange = ComboOnChange
          TagDispatch = 0
          VideString = '<<Aucun>>'
          DataType = 'GCTARIFPERIODE1'
        end
      end
      object GBA: TGroupBox
        Left = 278
        Top = 60
        Width = 271
        Height = 213
        Caption = 'A mettre à jour '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object TCOEF: THLabel
          Left = 4
          Top = 114
          Width = 105
          Height = 13
          Caption = 'Coefficient à appliquer'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TARRONDI: THLabel
          Left = 4
          Top = 142
          Width = 33
          Height = 13
          Caption = 'Arrondi'
          FocusControl = ARRONDIC
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TPXSAISI: THLabel
          Left = 4
          Top = 166
          Width = 72
          Height = 13
          Caption = 'Prix à appliquer'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object GBTARIFMAJ: TGroupBox
          Left = 4
          Top = 13
          Width = 263
          Height = 92
          Caption = 'Tarifs à mettre à jour '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          object TTYPETARIFDEST: THLabel
            Left = 4
            Top = 30
            Width = 64
            Height = 13
            Caption = 'Type de tarifs'
            FocusControl = TYPETARIFDEST
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object TPERIODEDEST: THLabel
            Left = 4
            Top = 57
            Width = 98
            Height = 13
            Caption = 'Période d'#39'application'
            FocusControl = PERIODEDEST
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object TYPETARIFDEST: THValComboBox
            Left = 113
            Top = 27
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
            OnChange = ComboOnChange
            TagDispatch = 0
            VideString = '<<Aucun>>'
            DataType = 'GCTARIFTYPE1'
          end
          object PERIODEDEST: THValComboBox
            Left = 113
            Top = 53
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
            OnChange = ComboOnChange
            TagDispatch = 0
            VideString = '<<Aucun>>'
            DataType = 'GCTARIFPERIODE1'
          end
        end
        object COEF: THNumEdit
          Left = 117
          Top = 112
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
          TabOrder = 1
          UseRounding = True
          Value = 1
          Validate = False
        end
        object ARRONDIC: THValComboBox
          Left = 117
          Top = 138
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'GCCODEARRONDI'
        end
        object PXSAISI: THNumEdit
          Left = 117
          Top = 164
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
          TabOrder = 3
          UseRounding = True
          Value = 1
          Validate = False
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 554
        Height = 272
      end
      object BFLECHEDROITE: TToolbarButton97
        Left = 261
        Top = 64
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
        OnClick = ClickFlecheDroite
      end
      object BFLECHEGAUCHE: TToolbarButton97
        Left = 261
        Top = 96
        Width = 28
        Height = 29
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
        OnClick = ClickFlecheGauche
      end
      object TLISTEET: THLabel
        Left = 20
        Top = 12
        Width = 114
        Height = 13
        Caption = 'Liste des établissements'
      end
      object TMAJ: THLabel
        Left = 301
        Top = 12
        Width = 135
        Height = 13
        Caption = 'Etablissement à mettre à jour'
      end
      object BFELCHETOUS: TToolbarButton97
        Left = 261
        Top = 136
        Width = 28
        Height = 29
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
        Left = 261
        Top = 172
        Width = 28
        Height = 29
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
        Left = 16
        Top = 29
        Width = 233
        Height = 234
        TabStop = False
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        FixedRows = 0
        TabOrder = 0
        OnDblClick = ClickFlecheDroite
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
      object GMAJ: THGrid
        Left = 301
        Top = 29
        Width = 233
        Height = 234
        TabStop = False
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        FixedRows = 0
        TabOrder = 1
        OnDblClick = ClickFlecheGauche
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
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object Bevel3: TBevel
        Left = 0
        Top = 0
        Width = 553
        Height = 272
      end
      object TRecap: THLabel
        Left = 6
        Top = 84
        Width = 59
        Height = 13
        Caption = 'Récapitulatif'
      end
      object PanelFin: TPanel
        Left = 6
        Top = 4
        Width = 540
        Height = 77
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 23
          Top = 8
          Width = 412
          Height = 26
          Caption = 
            'Le paramétrage est maintenant correctement renseigné pour permet' +
            'tre le lancement de la mise à jour des tarifs.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 41
          Width = 418
          Height = 26
          Caption = 
            'Si vous désirez revoir le paramétrage, il suffit de cliquer sur ' +
            'le bouton Précédent sinon, le bouton Fin, permet de débuter le t' +
            'raitement.'
          WordWrap = True
        end
      end
      object ListRecap: TListBox
        Left = 6
        Top = 100
        Width = 540
        Height = 167
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
      object BFinRecap: TButton
        Left = 444
        Top = 236
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
    Left = 2
    Top = 26
    Width = 159
    Height = 255
    inherited Image: TToolbarButton97
      Top = 38
      Width = 157
      Height = 187
    end
  end
  inherited cControls: TListBox
    Top = 262
  end
  object BStop: TButton [12]
    Left = 173
    Top = 326
    Width = 121
    Height = 23
    Caption = 'Arrêter le traitement'
    TabOrder = 5
    Visible = False
    OnClick = ClickBStop
  end
  object HMsgErr: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Il doit y avoir au moins un établissement à mettre à jour'
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
      'ATTENTION : Tarifs non enregistré(s) !'
      
        'ATTENTION : Ce tarif, en cours de traitement par un autre utilis' +
        'ateur, n'#39'a pas été enregistrée !'
      'enregistrement(s) modifié(s)'
      'Etablissement(s) mise à jour: '
      'Mise à jour réussie pour:'
      'Anomlie pour:')
    Left = 132
    Top = 16
  end
end

inherited FAssistTarifArt: TFAssistTarifArt
  Left = 223
  Top = 182
  Caption = 'Assistant de mise '#224' jour des bases tarifaires'
  ClientHeight = 334
  ClientWidth = 547
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 311
  end
  inherited lAide: THLabel
    Top = 281
    Width = 359
  end
  inherited bPrecedent: TToolbarButton97
    Left = 280
    Top = 305
  end
  inherited bSuivant: TToolbarButton97
    Left = 363
    Top = 305
  end
  inherited bFin: TToolbarButton97
    Left = 447
    Top = 305
  end
  inherited bAnnuler: TToolbarButton97
    Left = 197
    Top = 305
  end
  inherited bAide: TToolbarButton97
    Left = 114
    Top = 305
  end
  inherited Plan: THPanel
    Left = 180
  end
  inherited GroupBox1: THGroupBox
    Top = 293
    Width = 550
  end
  inherited P: THPageControl2
    Top = 1
    Width = 361
    Height = 296
    ActivePage = TabSheet3
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object TINTRO: THLabel
        Left = 6
        Top = 54
        Width = 312
        Height = 45
        AutoSize = False
        Caption = 
          'Cet assistant vous guide afin de param'#232'trer la mise '#224' jour des b' +
          'ases tarifaires des articles pr'#233'alablement s'#233'lectionn'#233's.'
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 353
        Height = 41
        Align = alTop
        Caption = 'Mise '#224' jour des bases tarifaires'
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
      object GBParamBase: TGroupBox
        Left = 0
        Top = 104
        Width = 338
        Height = 125
        Caption = 'Modification du prix d'#39'achat de base'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TARRONDIBASE: THLabel
          Left = 15
          Top = 89
          Width = 85
          Height = 13
          Caption = '&M'#233'thode d'#39'arrondi'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TTYPEMAJPRIXBASE: THLabel
          Left = 15
          Top = 40
          Width = 92
          Height = 13
          Caption = '&Type de mise '#224' jour'
          FocusControl = TYPEMAJPRIXBASE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TVALEURPRIXBASE: THLabel
          Left = 15
          Top = 64
          Width = 30
          Height = 13
          Caption = '&Valeur'
          FocusControl = VALEURPRIXBASE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object ARRONDIBASE: THValComboBox
          Left = 128
          Top = 85
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCCODEARRONDI'
        end
        object TYPEMAJPRIXBASE: THValComboBox
          Left = 128
          Top = 36
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 0
          OnChange = TYPEMAJPRIXBASEChange
          TagDispatch = 0
          Plus = 'AND CO_CODE Like "P%" AND CO_CODE<> "PO1"'
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCMAJTARIF'
        end
        object VALEURPRIXBASE: THNumEdit
          Left = 128
          Top = 60
          Width = 145
          Height = 21
          TabStop = False
          Color = clBtnFace
          Decimals = 2
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object GBParamTTC: TGroupBox
        Left = 0
        Top = 99
        Width = 338
        Height = 88
        Caption = 'Modification du prix TTC'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TARRONDITTC: THLabel
          Left = 15
          Top = 62
          Width = 85
          Height = 13
          Caption = 'M'#233'thode d'#39'&arrondi'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TTYPEMAJPRIXTTC: THLabel
          Left = 15
          Top = 16
          Width = 92
          Height = 13
          Caption = 'Type de mise '#224' &jour'
          FocusControl = TYPEMAJPRIXTTC
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TVALEURPRIXTTC: THLabel
          Left = 15
          Top = 39
          Width = 30
          Height = 13
          Caption = 'Va&leur'
          FocusControl = VALEURPRIXTTC
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object ARRONDITTC: THValComboBox
          Left = 128
          Top = 58
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCCODEARRONDI'
        end
        object TYPEMAJPRIXTTC: THValComboBox
          Left = 128
          Top = 12
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 0
          OnChange = TYPEMAJPRIXTTCChange
          TagDispatch = 0
          Plus = 'AND CO_CODE Like "P%" AND CO_CODE<> "PO1"'
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCMAJTARIF'
        end
        object VALEURPRIXTTC: THNumEdit
          Left = 128
          Top = 35
          Width = 145
          Height = 21
          TabStop = False
          Color = clBtnFace
          Decimals = 2
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
      end
      object GBParamHT: TGroupBox
        Left = 0
        Top = 5
        Width = 338
        Height = 88
        Caption = 'Modification du prix HT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TARRONDIHT: THLabel
          Left = 15
          Top = 62
          Width = 85
          Height = 13
          Caption = '&M'#233'thode d'#39'arrondi'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TTYPEMAJPRIXHT: THLabel
          Left = 15
          Top = 17
          Width = 92
          Height = 13
          Caption = '&Type de mise '#224' jour'
          FocusControl = TYPEMAJPRIXHT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TVALEURPRIXHT: THLabel
          Left = 15
          Top = 40
          Width = 30
          Height = 13
          Caption = '&Valeur'
          FocusControl = VALEURPRIXHT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object ARRONDIHT: THValComboBox
          Left = 128
          Top = 58
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCCODEARRONDI'
        end
        object TYPEMAJPRIXHT: THValComboBox
          Left = 128
          Top = 13
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 0
          OnChange = TYPEMAJPRIXHTChange
          TagDispatch = 0
          Plus = 'AND CO_CODE Like "P%" AND CO_CODE<> "PO1"'
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCMAJTARIF'
        end
        object VALEURPRIXHT: THNumEdit
          Left = 128
          Top = 36
          Width = 145
          Height = 21
          TabStop = False
          Color = clBtnFace
          Decimals = 2
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
      end
      object GroupBox2: TGroupBox
        Left = 0
        Top = 190
        Width = 338
        Height = 69
        Caption = 'Calculer les prix de vente'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object Label1: TLabel
          Left = 320
          Top = 42
          Width = 8
          Height = 13
          Caption = '%'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CALCPASAUTOPV: TCheckBox
          Left = 128
          Top = 40
          Width = 145
          Height = 17
          Caption = 'Si diff'#233'rence sup'#233'rieure '#224
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object HT: TRadioButton
          Left = 12
          Top = 17
          Width = 37
          Height = 17
          Caption = 'HT'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TabStop = True
        end
        object TTC: TRadioButton
          Left = 52
          Top = 17
          Width = 41
          Height = 17
          Caption = 'TTC'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object CALCAUTOPV: TCheckBox
          Left = 128
          Top = 17
          Width = 145
          Height = 17
          Caption = 'Si calcul en automatique'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object ECART: THNumEdit
          Left = 276
          Top = 38
          Width = 37
          Height = 21
          Decimals = 2
          Digits = 4
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          TabOrder = 4
          UseRounding = True
          Validate = False
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object TRecap: THLabel
        Left = 5
        Top = 108
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object ListRecap: TListBox
        Left = 5
        Top = 126
        Width = 329
        Height = 116
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 0
      end
      object PanelFin: TPanel
        Left = 5
        Top = 4
        Width = 329
        Height = 97
        TabOrder = 1
        object TTextFin1: THLabel
          Left = 23
          Top = 8
          Width = 279
          Height = 26
          Caption = 
            'Le param'#232'trage est correctement renseign'#233' pour permettre le lanc' +
            'ement de la  mise '#224' jour des bases tarifaires.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 49
          Width = 285
          Height = 39
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon, le bouton Fin, permet de d'#233'buter le t' +
            'raitement.'
          WordWrap = True
        end
      end
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      'ATTENTION : Tarif non enregistr'#233' !'
      
        'ATTENTION : Ce tarif, en cours de traitement par un autre utilis' +
        'ateur, n'#39'a pas '#233't'#233' enregistr'#233'e !'
      'enregistrement(s) modifi'#233#39's)'
      '')
  end
end

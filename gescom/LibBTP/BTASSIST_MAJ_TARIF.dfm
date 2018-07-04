inherited BTASSIST_MAJ_TARIF: TBTASSIST_MAJ_TARIF
  Left = 365
  Top = 270
  Caption = 'Assistant de mise '#224' jour des bases tarifaires articles'
  PixelsPerInch = 96
  TextHeight = 13
  inherited P: THPageControl2
    ActivePage = TSDEPART
    object TSDEPART: TTabSheet
      Caption = 'TSDEPART'
      object RBMAJAPA: TRadioButton
        Left = 52
        Top = 134
        Width = 205
        Height = 17
        Caption = 'Mise '#224' jour des prix d'#39'achat'
        TabOrder = 0
        OnClick = RBMAJAPAClick
      end
      object RBMAJPV: TRadioButton
        Left = 52
        Top = 72
        Width = 205
        Height = 17
        Caption = 'Mise '#224' jour des prix de vente'
        Checked = True
        TabOrder = 1
        TabStop = True
        OnClick = RBMAJPVClick
      end
    end
    object TSACHAT: TTabSheet
      Caption = 'TSACHAT'
      ImageIndex = 1
      object HLabel1: THLabel
        Left = 27
        Top = 76
        Width = 61
        Height = 13
        Caption = 'Pourcentage'
      end
      object Label1: TLabel
        Left = 0
        Top = 0
        Width = 338
        Height = 20
        Align = alTop
        Alignment = taCenter
        Caption = 'Mise '#224' jour des prix d'#39'achats'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object POURCENTACH: THNumEdit
        Left = 104
        Top = 72
        Width = 55
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object AUGMENTACH: TRadioButton
        Left = 197
        Top = 64
        Width = 113
        Height = 17
        Caption = 'Augmentation'
        Checked = True
        TabOrder = 1
        TabStop = True
      end
      object DIMINUACH: TRadioButton
        Left = 197
        Top = 81
        Width = 113
        Height = 17
        Caption = 'Diminution'
        TabOrder = 2
      end
      object SURPVOK: TCheckBox
        Left = 28
        Top = 124
        Width = 273
        Height = 17
        Caption = 'R'#233'percution sur le prix de vente'
        TabOrder = 3
      end
    end
    object TSVENTE: TTabSheet
      Caption = 'TSVENTE'
      ImageIndex = 2
      object Label2: TLabel
        Left = 0
        Top = 0
        Width = 338
        Height = 20
        Align = alTop
        Alignment = taCenter
        Caption = 'Mise '#224' jour des prix de vente'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 27
        Top = 76
        Width = 61
        Height = 13
        Caption = 'Pourcentage'
      end
      object POURCENTVTE: THNumEdit
        Left = 104
        Top = 72
        Width = 55
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object AUGMENTVTE: TRadioButton
        Left = 197
        Top = 64
        Width = 113
        Height = 17
        Caption = 'Augmentation'
        Checked = True
        TabOrder = 1
        TabStop = True
      end
      object DIMINUVTE: TRadioButton
        Left = 197
        Top = 81
        Width = 113
        Height = 17
        Caption = 'Diminution'
        TabOrder = 2
      end
    end
    object TSRECAP: TTabSheet
      Caption = 'TSRECAP'
      ImageIndex = 3
      object TRecap: THLabel
        Left = 5
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
end

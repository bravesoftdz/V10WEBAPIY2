inherited FAssistMajCptSouche: TFAssistMajCptSouche
  Left = 147
  Top = 189
  Caption = 'Assistant Recalcul des compteurs'
  PixelsPerInch = 96
  TextHeight = 13
  inherited bPrecedent: TToolbarButton97
    Left = 306
    Width = 70
  end
  inherited bSuivant: TToolbarButton97
    Left = 381
    Width = 70
  end
  inherited bFin: TToolbarButton97
    Left = 457
    Width = 70
  end
  inherited bAnnuler: TToolbarButton97
    Left = 232
    Width = 70
  end
  inherited bAide: TToolbarButton97
    Left = 157
    Width = 70
  end
  object bImprimer: TToolbarButton97 [7]
    Left = 81
    Top = 308
    Width = 70
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = 'Imprimer'
    Flat = False
    OnClick = bImprimerClick
  end
  inherited P: THPageControl2
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object LExplication: THLabel
        Left = 8
        Top = 54
        Width = 317
        Height = 41
        AutoSize = False
        Caption = 
          'Le traitement va positionner les compteurs de souche au plus gra' +
          'nd des num'#233'ros utilis'#233's dans les documents pour la nature de pi'#232 +
          'ce s'#233'lectionn'#233'e'
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 338
        Height = 41
        Align = alTop
        Caption = 'Recalcul des compteurs'
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
      object GBParam: TGroupBox
        Left = 0
        Top = 147
        Width = 338
        Height = 89
        Caption = 'Param'#232'trage'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TNaturePieceG: THLabel
          Left = 21
          Top = 34
          Width = 76
          Height = 13
          Caption = 'Nature de pi'#232'ce'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object CBNaturePieceG: THValComboBox
          Left = 107
          Top = 30
          Width = 142
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
          OnChange = CBNaturePieceGChange
          TagDispatch = 0
          Vide = True
          VideString = '<<Toutes>>'
          DataType = 'GCNATUREPIECEG'
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object TRecap: THLabel
        Left = 5
        Top = 84
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object PanelFin: TPanel
        Left = 5
        Top = 4
        Width = 329
        Height = 77
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 23
          Top = 8
          Width = 279
          Height = 26
          Caption = 
            'Le param'#232'trage est correctement renseign'#233' pour permettre le lanc' +
            'ement de la mise '#224' jour des compteurs.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 34
          Width = 285
          Height = 39
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon, le bouton Fin, permet de d'#233'buter le t' +
            'raitement.'
          WordWrap = True
        end
      end
      object LBRecap: TListBox
        Left = 5
        Top = 99
        Width = 329
        Height = 142
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
    end
  end
end

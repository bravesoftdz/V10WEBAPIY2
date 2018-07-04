inherited FAssistStock: TFAssistStock
  Left = 381
  Top = 234
  Caption = 'Assistant d'#39'affectation'
  PixelsPerInch = 96
  TextHeight = 13
  inherited bAide: TToolbarButton97
    Visible = True
  end
  inherited P: TPageControl
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object PBevel1: TBevel
        Left = 0
        Top = 41
        Width = 338
        Height = 203
        Align = alClient
      end
      object TINTRO: THLabel
        Left = 22
        Top = 52
        Width = 296
        Height = 45
        AutoSize = False
        Caption = 
          '   Cet assistant vous guide afin de paramètrer l'#39'affectation aut' +
          'omatique de votre stock aux différente commandes clients en cour' +
          's.'
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 338
        Height = 41
        Align = alTop
        Caption = 'Assistant d'#39'affectation de stock'
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
      object GBPRIOR: TGroupBox
        Left = 17
        Top = 104
        Width = 310
        Height = 113
        Caption = 'Critères de sélection'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TPRIORITE: THLabel
          Left = 10
          Top = 25
          Width = 32
          Height = 13
          Caption = '&Priorité'
          Color = clBtnFace
          FocusControl = PRIORITE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TDATELIVRMAX: TLabel
          Left = 10
          Top = 56
          Width = 89
          Height = 13
          Caption = '&Date livraison Maxi'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object PRIORITE: THValComboBox
          Left = 112
          Top = 21
          Width = 101
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
          Items.Strings = (
            'Date Pièce'
            'Date livraison'
            'Total H.T.')
          TagDispatch = 0
          Values.Strings = (
            'GP_DATEPIECE'
            'GL_DATELIVRAISON'
            'GP_TOTALHT')
        end
        object CBASCENDANT: TCheckBox
          Left = 224
          Top = 23
          Width = 77
          Height = 17
          Caption = '&Ascendant'
          Checked = True
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          State = cbChecked
          TabOrder = 1
        end
        object DATELIVRMAX: THCritMaskEdit
          Left = 112
          Top = 52
          Width = 101
          Height = 21
          TabStop = False
          EditMask = '!99 >L<LL 0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 11
          ParentFont = False
          TabOrder = 2
          Text = '           '
          OnExit = DATELIVRMAXExit
          TagDispatch = 0
          OpeType = otDate
          DefaultDate = od2099
          ControlerDate = True
        end
        object CBRELIQUATLIG: TCheckBox
          Left = 11
          Top = 87
          Width = 178
          Height = 17
          Caption = 'Accepter les &reliquats à la ligne'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 3
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object PBevel2: TBevel
        Left = 0
        Top = 0
        Width = 338
        Height = 244
        Align = alClient
      end
      object TRecap: THLabel
        Left = 18
        Top = 108
        Width = 59
        Height = 13
        Caption = 'Récapitulatif'
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
            'Le paramètrage est maintenant correctement renseigné pour permet' +
            'tre le lancement de l'#39'affectation de stock'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 49
          Width = 285
          Height = 39
          Caption = 
            'Si vous désirez revoir le paramétrage, il suffit de cliquer sur ' +
            'le bouton Précédent sinon, le bouton Fin, permet de débuter le t' +
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
  object HStk: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Oui'
      'Non'
      'Ascendant'
      'Descendant')
    Left = 136
    Top = 8
  end
  object HMsgErr: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'n'#39'est pas une date valide.')
    Left = 56
    Top = 4
  end
end

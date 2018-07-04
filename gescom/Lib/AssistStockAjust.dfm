inherited FStockAjust: TFStockAjust
  HelpContext = 110000321
  Caption = 'Assistant de r'#233'ajustement des stocks'
  PixelsPerInch = 96
  TextHeight = 13
  inherited bFin: TToolbarButton97
    Enabled = False
  end
  inherited P: THPageControl2
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object TWarning: THLabel
        Left = 7
        Top = 59
        Width = 318
        Height = 26
        AutoSize = False
        Caption = 'Attention ce traitement peut '#234'tre long'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object HLabel2: THLabel
        Left = 8
        Top = 84
        Width = 317
        Height = 47
        AutoSize = False
        Caption = 
          'Le traitement va v'#233'rifier les quantit'#233's de stock et les prix de ' +
          'revalorisation en scrutant les mouvements effectu'#233's depuis le de' +
          'rnier traitement de fin de p'#233'riode.'
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 338
        Height = 41
        Align = alTop
        Caption = 'R'#233'ajustement des stocks'
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
        object TDepot: THLabel
          Left = 29
          Top = 34
          Width = 29
          Height = 13
          Caption = 'D'#233'p'#244't'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TCodeArticle: THLabel
          Left = 29
          Top = 60
          Width = 29
          Height = 13
          Caption = 'Article'
          FocusControl = CodeArticle
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          Visible = False
        end
        object Depot: THValComboBox
          Left = 67
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
          TagDispatch = 0
          Vide = True
          DataType = 'GCDEPOT'
        end
        object CodeArticle: THCritMaskEdit
          Left = 67
          Top = 56
          Width = 142
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          Visible = False
          TagDispatch = 0
          Operateur = Egal
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object HLabel1: THLabel
        Left = 11
        Top = 60
        Width = 312
        Height = 45
        AutoSize = False
        WordWrap = True
      end
      object TRecap: THLabel
        Left = 5
        Top = 105
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object TProgressBar: THLabel
        Left = 5
        Top = 207
        Width = 329
        Height = 14
        AutoSize = False
        WordWrap = True
      end
      object TArticle: THLabel
        Left = 5
        Top = 245
        Width = 329
        Height = 18
        AutoSize = False
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
            'ement de la v'#233'rification des stocks.'
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
      object LBX_Param: TListBox
        Left = 5
        Top = 123
        Width = 329
        Height = 73
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
      object ProgressBar: TProgressBar
        Left = 5
        Top = 223
        Width = 329
        Height = 18
        Step = 1
        TabOrder = 2
        Visible = False
      end
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      
        '1;?caption?;Voulez-vous quitter l'#39'assistant et annuler l'#39'op'#233'rati' +
        'on?;Q;YN;Y;C;'
      
        '2;?caption?;Probl'#232'mes rencontr'#233's : l'#39'op'#233'ration a '#233't'#233' annul'#233'e;W;O' +
        ';O;O;'
      '3;?caption?;Aucun r'#233'ajustement n'#39'est n'#233'cessaire;E;O;O;O;'
      '4;?caption?;Le traitement s'#39'est correctement effectu'#233';E;O;O;O;'
      'V'#233'rification du d'#233'p'#244't :  '
      'Article :  ')
    Left = 75
    Top = 8
  end
end

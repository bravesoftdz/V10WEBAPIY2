inherited FCmdeContreM: TFCmdeContreM
  Left = 113
  Top = 193
  Caption = 'Assistant de g'#233'n'#233'ration de commande'
  ClientWidth = 532
  PixelsPerInch = 96
  TextHeight = 13
  inherited bPrecedent: TToolbarButton97
    Left = 285
  end
  inherited bSuivant: TToolbarButton97
    Left = 360
  end
  inherited bFin: TToolbarButton97
    Left = 448
  end
  inherited bAnnuler: TToolbarButton97
    Left = 198
  end
  inherited bAide: TToolbarButton97
    Left = 115
  end
  inherited P: THPageControl2
    Width = 356
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object TINTRO: THLabel
        Left = 22
        Top = 55
        Width = 296
        Height = 41
        AutoSize = False
        Caption = 
          '   Cet assistant vous guide dans la g'#233'n'#233'ration automatique de vo' +
          's commandes fournisseurs en fonction de votre s'#233'lection initiale' +
          '.'
        WordWrap = True
      end
      object GBParam: TGroupBox
        Left = 3
        Top = 104
        Width = 338
        Height = 105
        Caption = 'Param'#232'trage'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel1: THLabel
          Left = 18
          Top = 31
          Width = 79
          Height = 13
          Caption = 'D'#233'lai de s'#233'curit'#233
          FocusControl = NbJourSecu
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel2: THLabel
          Left = 155
          Top = 31
          Width = 28
          Height = 13
          Caption = 'jour(s)'
          FocusControl = NbJourSecu
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object NbJourSecu: THNumEdit
          Left = 107
          Top = 27
          Width = 41
          Height = 21
          TabOrder = 0
          Decimals = 0
          Digits = 12
          Masks.PositiveMask = '#0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object CBRECUPBLOCNOTE: TCheckBox
          Left = 17
          Top = 61
          Width = 285
          Height = 17
          Caption = 'R'#233'cup'#233'ration du bloc-note de la ligne de pi'#232'ce'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
        end
        object CBECLATCDEFOU: TCheckBox
          Left = 17
          Top = 84
          Width = 305
          Height = 17
          Caption = 'Eclatement des cdes fournisseurs par chantier et destination'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 2
          Visible = False
        end
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 348
        Height = 41
        Align = alTop
        Caption = 'G'#233'n'#233'ration commandes fournisseurs'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object TRecap: THLabel
        Left = 18
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
          Width = 287
          Height = 26
          Caption = 
            'Le param'#232'trage est maintenant correctement renseign'#233' pour permet' +
            'tre la g'#233'n'#233'ration des commandes d'#39'achat.'
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
end

inherited FAssistFinInv: TFAssistFinInv
  Left = 123
  Top = 173
  HelpContext = 300000454
  Caption = 'Assistant de fin d'#39'inventaire'
  ClientHeight = 332
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Caption = 'Etape 1/1'
  end
  inherited bFin: TToolbarButton97
    Caption = '&Exécuter'
  end
  inherited bAide: TToolbarButton97
    Visible = True
  end
  inherited P: TPageControl
    Height = 248
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
    end
  end
  object Panel1: TPanel [12]
    Left = 180
    Top = 8
    Width = 345
    Height = 273
    TabOrder = 5
    object Bevel1: TBevel
      Left = 4
      Top = 44
      Width = 337
      Height = 225
    end
    object HLabel1: THLabel
      Left = 16
      Top = 52
      Width = 311
      Height = 39
      Caption = 
        'Attention, cette procédure termine votre inventaire, en supprima' +
        'nt définitivement vos listes et inventaires transmis pour tous l' +
        'es établissements sélectionnés.'
      WordWrap = True
    end
    object HLabel2: THLabel
      Left = 16
      Top = 93
      Width = 308
      Height = 26
      Caption = 
        'Assurez-vous d'#39'avoir bien édité tous vos états d'#39'inventaire, ava' +
        'nt de la lancer.'
      WordWrap = True
    end
    object PTITRE: THPanel
      Left = 1
      Top = 1
      Width = 343
      Height = 41
      Align = alTop
      Caption = 'Fin d'#39'inventaire'
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
    object GBCRITERE: TGroupBox
      Left = 12
      Top = 126
      Width = 325
      Height = 139
      Caption = 'Critères de sélection'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object TETABLISS: THLabel
        Left = 10
        Top = 25
        Width = 70
        Height = 13
        Caption = 'Etablissements'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object CB_SUPPLISTE: TCheckBox
        Left = 11
        Top = 66
        Width = 190
        Height = 17
        Caption = 'Suppression des listes d'#39'inventaire'
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
        TabOrder = 0
      end
      object CB_INVTRANS: TCheckBox
        Left = 11
        Top = 87
        Width = 198
        Height = 17
        Caption = 'Suppression des inventaires transmis'
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
      object ETABLISSEMENTS: THMultiValComboBox
        Left = 96
        Top = 21
        Width = 225
        Height = 21
        TabOrder = 2
        Text = '<<Tous>>'
        Abrege = False
        DataType = 'GCDEPOT'
        Complete = False
        OuInclusif = False
      end
      object RG_TYPEDOC: THRadioGroup
        Left = 215
        Top = 51
        Width = 106
        Height = 79
        Caption = 'Type de document'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBtnText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemIndex = 1
        Items.Strings = (
          'Tous'
          'Validé'
          'Non validé')
        ParentFont = False
        TabOrder = 3
        Abrege = False
        Vide = False
        Values.Strings = (
          '0'
          '1'
          '2')
      end
    end
  end
end

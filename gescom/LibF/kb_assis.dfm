inherited FCopieClavierSoc: TFCopieClavierSoc
  HelpContext = 113000300
  Caption = 'Copie pavé depuis société '
  ClientWidth = 550
  PixelsPerInch = 96
  TextHeight = 13
  inherited bAide: TToolbarButton97
    Visible = True
  end
  inherited P: TPageControl
    ActivePage = pCaisse
    object PSoc: TTabSheet
      Caption = 'PSoc'
      object Label6: TLabel
        Left = 45
        Top = 101
        Width = 71
        Height = 13
        Caption = 'Société source'
      end
      object Label12: TLabel
        Left = 31
        Top = 155
        Width = 281
        Height = 75
        AutoSize = False
        Caption = 
          'Choisissez ci-dessus la société à partir de laquelle vous souhai' +
          'tez recopier le pavé. '
        WordWrap = True
      end
      object Label3: TLabel
        Left = 109
        Top = 36
        Width = 119
        Height = 13
        Caption = 'Initialisation du pavé'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel2: TBevel
        Left = 109
        Top = 51
        Width = 119
        Height = 5
        Shape = bsBottomLine
      end
      object FSocSource: THValComboBox
        Left = 124
        Top = 97
        Width = 162
        Height = 21
        Hint = 
          'La société source est la société qui contient les paramètres à r' +
          'ecopier'
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FSocSourceChange
        TagDispatch = 0
      end
    end
    object pCaisse: TTabSheet
      Caption = 'pCaisse'
      ImageIndex = 1
      object pnlCaisse1: TPanel
        Left = 0
        Top = 0
        Width = 338
        Height = 244
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        object Label1: TLabel
          Left = 109
          Top = 36
          Width = 119
          Height = 13
          Caption = 'Initialisation du pavé'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object Bevel1: TBevel
          Left = 109
          Top = 51
          Width = 119
          Height = 5
          Shape = bsBottomLine
        end
        object TCbCaisse: TLabel
          Left = 45
          Top = 101
          Width = 31
          Height = 13
          Caption = 'Caisse'
          FocusControl = cbCaisse
        end
        object Label4: TLabel
          Left = 31
          Top = 155
          Width = 281
          Height = 75
          AutoSize = False
          Caption = 
            'Choisissez ci-dessus la caisse à partir de laquelle vous souhait' +
            'ez recopier le pavé. '
          WordWrap = True
        end
        object cbCaisse: THValComboBox
          Left = 124
          Top = 97
          Width = 162
          Height = 21
          Hint = 
            'La société source est la société qui contient les paramètres à r' +
            'ecopier'
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 0
          OnChange = cbCaisseChange
          TagDispatch = 0
        end
      end
      object PnlCaisse2: TPanel
        Left = 0
        Top = 0
        Width = 338
        Height = 244
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 1
        object LblAvis: TFlashingLabel
          Left = 74
          Top = 115
          Width = 189
          Height = 13
          Caption = 'Recherches des caisses paramètres......'
          Color = clActiveCaption
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          Flashing = True
        end
      end
    end
    object pTypes: TTabSheet
      Caption = 'pTypes'
      ImageIndex = 2
      object Label13: TLabel
        Left = 4
        Top = 228
        Width = 303
        Height = 13
        Caption = 'Cochez les types de boutons à recopier dans la caisse en cours.'
        Transparent = True
      end
      object grTypes: TGroupBox
        Left = 2
        Top = 3
        Width = 332
        Height = 220
        Caption = ' Types de boutons '
        TabOrder = 0
        object CheckBox1: TCheckBox
          Left = 10
          Top = 20
          Width = 97
          Height = 17
          Caption = 'CheckBox1'
          TabOrder = 0
          OnClick = clickCheck
        end
        object CheckBox2: TCheckBox
          Left = 10
          Top = 40
          Width = 97
          Height = 17
          Caption = 'CheckBox1'
          TabOrder = 1
        end
      end
    end
  end
  object DBSource: TDatabase
    DatabaseName = 'DBSource'
    SessionName = 'Default'
    Left = 62
    Top = 248
  end
end

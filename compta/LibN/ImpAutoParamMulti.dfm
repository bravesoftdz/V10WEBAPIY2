inherited FImpAutoParamMulti: TFImpAutoParamMulti
  Left = 308
  Top = 233
  Height = 173
  Caption = 'Paramétrage de l'#39'import automatique'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 111
  end
  object HPanel1: THPanel [1]
    Left = 0
    Top = 0
    Width = 317
    Height = 111
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HLabel1: THLabel
      Left = 16
      Top = 50
      Width = 32
      Height = 13
      Caption = 'Format'
    end
    object HLabel2: THLabel
      Left = 16
      Top = 82
      Width = 81
      Height = 13
      Caption = 'Scénario d'#39'import'
    end
    object HLabel3: THLabel
      Left = 16
      Top = 16
      Width = 236
      Height = 13
      Caption = 'Veuillez choisir ici le format et le scénario d'#39'import :'
    end
    object BSCENARIO: TToolbarButton97
      Left = 284
      Top = 77
      Width = 23
      Height = 22
      Hint = 'Accéder au scénario'
      Glyph.Data = {
        42010000424D4201000000000000760000002800000011000000110000000100
        040000000000CC00000000000000000000001000000010000000000000000000
        BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
        77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
        F0F0FF0F0F07700000007000000000000007700000007770EEEEEEE077777000
        00007770E0000E077777700000007770E070E0077777700000007770E00EE700
        7777700000007770E0EE07000777700000007770EEE0770BB077700000007770
        EE07770BB077700000007770E0777770BB0770000000777007777770B0077000
        00007770777777770BB070000000777777777777000770000000777777777777
        777770000000}
      OnClick = BSCENARIOClick
    end
    object SCENARIOMULTI: THCritMaskEdit
      Left = 112
      Top = 78
      Width = 165
      Height = 21
      TabOrder = 1
      TagDispatch = 0
      DataType = 'CPSCENARIO'
      ElipsisButton = True
    end
    object FORMATMULTI: THValComboBox
      Left = 112
      Top = 46
      Width = 165
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FORMATMULTIChange
      TagDispatch = 0
      DataType = 'TTFORMATECRITURE'
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 148
    Top = 112
  end
end

inherited FAssistPLC: TFAssistPLC
  Left = 239
  Top = 160
  Caption = 'Chargement d'#39'un dossier type'
  KeyPreview = True
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Visible = False
  end
  object HLabel1: THLabel [2]
    Left = 187
    Top = 187
    Width = 152
    Height = 13
    Caption = 'Longueur des comptes géréraux'
    Contour = False
    Ombre = False
    OmbreDecalX = 0
    OmbreDecalY = 0
    OmbreColor = clBlack
    Tag2 = 0
    Tag3 = 0
    TagExport = 0
  end
  object BtSupp: TToolbarButton97 [3]
    Left = 316
    Top = 297
    Width = 23
    Height = 22
    Hint = 'Suppression plan type'
    Glyph.Data = {
      66010000424D6601000000000000760000002800000014000000140000000100
      040000000000F000000000000000000000001000000010000000000000000000
      8000008000000080800080000000800080008080000080808000C0C0C0000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00888888888888
      8888888800008888888888888888889800008898888888888888898800008899
      88777777777798880000889990000000000998880000888990BFFFBFFF998888
      0000888899FCCCCCCF97888800008888999FBFFFB9978888000088888999CCC9
      990788880000888880999FB99F0788880000888880FC9999CF07888800008888
      80FF9999BF0788880000888880FC9999000788880000888880B99F099F078888
      0000888880999F099998888800008888999FBF0F089988880000889999000000
      8889988800008899988888888888898800008888888888888888889800008888
      88888888888888880000}
    OnClick = BtSuppClick
  end
  inherited bPrecedent: TButton
    Left = 236
    Width = 44
    Visible = False
  end
  inherited bSuivant: TButton
    Left = 188
    Width = 43
    Visible = False
  end
  inherited bFin: TButton
    Caption = '&Chargement'
  end
  inherited bAnnuler: TButton
    Left = 346
    Caption = '&Sortie'
  end
  inherited P: TPageControl
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object SO_LGCPTEGEN: TSpinEdit
        Left = 288
        Top = 208
        Width = 39
        Height = 22
        Color = clYellow
        MaxValue = 10
        MinValue = 6
        TabOrder = 0
        Value = 6
        Visible = False
      end
      object SO_LGCPTEAUX: TSpinEdit
        Left = 245
        Top = 206
        Width = 39
        Height = 22
        Color = clYellow
        MaxValue = 10
        MinValue = 6
        TabOrder = 1
        Value = 6
        Visible = False
      end
      object LSTMAQ: TListBox
        Left = 0
        Top = 0
        Width = 338
        Height = 244
        Align = alClient
        ItemHeight = 13
        TabOrder = 2
        OnClick = LSTMAQClick
        OnDblClick = LSTMAQDblClick
      end
    end
  end
end

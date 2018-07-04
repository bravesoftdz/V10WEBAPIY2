inherited AssitMaq: TAssitMaq
  Left = 291
  Top = 178
  Caption = 'Création et Modification d'#39'une maquette'
  ClientHeight = 324
  ClientWidth = 531
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Visible = False
  end
  inherited lAide: THLabel
    Left = 182
    Top = 264
    Caption = 'Selectionner une maquette'
    OmbreColor = clFuchsia
  end
  object Bimprime: TToolbarButton97 [2]
    Left = 228
    Top = 297
    Width = 23
    Height = 22
    Hint = 'Impression de la maquette'
    Glyph.Data = {
      F6000000424DF600000000000000760000002800000010000000100000000100
      0400000000008000000000000000000000001000000000000000000000000000
      80000080000000808000800000008000800080800000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FFFFFFFFFFFF
      FFFFFFFFFFFF000000FFFFF4FFF00000000FFFF44FF033333000444444F00000
      0030FFF44FF0F7F7F700FFF4FFF000000780FFFFFFFF0FFFF00FFF00000F0F88
      F0FFFF0FFF0FF0FFFF0F00000F0FFF00000F0FFF0F0FFFFFFFFF0F8F000FFFFF
      FFFF0F8F00FFFFFFFFFF0FF00FFFFFFFFFFF0000FFFFFFFFFFFF}
    OnClick = BimprimeClick
  end
  object BtSupp: TToolbarButton97 [3]
    Left = 196
    Top = 297
    Width = 23
    Height = 22
    Hint = 'Suppression de la maquette'
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
  inherited bPrecedent: TToolbarButton97
    Left = 312
    Width = 20
    Visible = False
  end
  inherited bSuivant: TToolbarButton97
    Left = 340
    Caption = '&Création'
  end
  inherited bFin: TToolbarButton97
    Caption = '&Modification'
  end
  inherited bAnnuler: TToolbarButton97
    Left = 258
  end
  inherited bAide: TToolbarButton97
    Left = 39
    Top = 293
  end
  inherited Plan: TPanel
    Top = 3
    Height = 244
  end
  inherited P: TPageControl
    Height = 235
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object LSTMAQ: TListBox
        Left = 0
        Top = 0
        Width = 338
        Height = 207
        Align = alClient
        ItemHeight = 13
        TabOrder = 0
        OnDblClick = LSTMAQDblClick
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object HLabel4: THLabel
        Left = 79
        Top = 21
        Width = 159
        Height = 16
        AutoSize = False
        Caption = 'Création d'#39'une maquette'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LabelNUMSTD: TLabel
        Left = 9
        Top = 68
        Width = 96
        Height = 13
        Caption = 'Numéro de standard'
      end
      object BTNUMSTD: TToolbarButton97
        Left = 150
        Top = 64
        Width = 23
        Height = 22
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC000000CE0E0000D80E00001000000000000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        OnClick = BTNUMSTDClick
      end
      object Label1: TLabel
        Left = 9
        Top = 105
        Width = 33
        Height = 13
        Caption = 'Libellé '
      end
      object GroupBox3: TGroupBox
        Left = 79
        Top = 34
        Width = 138
        Height = 7
        TabOrder = 0
      end
      object NUMSTD: TSpinEdit
        Tag = 64102
        Left = 114
        Top = 64
        Width = 39
        Height = 22
        MaxValue = 99
        MinValue = 21
        TabOrder = 1
        Value = 21
        OnChange = NUMSTDChange
      end
      object LIBSTD: TEdit
        Left = 114
        Top = 101
        Width = 245
        Height = 21
        TabOrder = 2
      end
    end
  end
end

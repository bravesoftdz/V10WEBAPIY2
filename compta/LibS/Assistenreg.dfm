inherited FASSISTENREG: TFASSISTENREG
  Left = 219
  Top = 165
  Caption = 'Enregistrement du dossier type'
  PixelsPerInch = 96
  TextHeight = 13
  inherited bFin: TToolbarButton97
    Caption = '&Enregistrer'
  end
  inherited P: TPageControl
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object LabelNUMSTD: TLabel
        Left = 32
        Top = 8
        Width = 96
        Height = 13
        Caption = 'Numéro de standard'
      end
      object Labelibelle: TLabel
        Left = 32
        Top = 36
        Width = 30
        Height = 13
        Caption = 'Libellé'
      end
      object BTNUMSTD: TToolbarButton97
        Left = 184
        Top = 4
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
      object NUMSTD: TSpinEdit
        Tag = 64102
        Left = 144
        Top = 4
        Width = 39
        Height = 22
        MaxValue = 99
        MinValue = 21
        TabOrder = 0
        Value = 21
        OnChange = NUMSTDChange
      end
      object LIBELLESTD: TEdit
        Left = 144
        Top = 28
        Width = 181
        Height = 21
        TabOrder = 1
      end
      object GD: THGrid
        Left = 8
        Top = 51
        Width = 320
        Height = 191
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 7
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        SortedCol = -1
        Titres.Strings = (
          ''
          '')
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
        ColWidths = (
          225
          90)
      end
    end
  end
end

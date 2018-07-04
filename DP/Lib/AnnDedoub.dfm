inherited FAnnDedoub: TFAnnDedoub
  Left = 321
  Top = 166
  Caption = 'Assistant d'#233'doublonnage annuaire'
  ClientWidth = 539
  PixelsPerInch = 96
  TextHeight = 13
  inherited lAide: THLabel
    Width = 351
  end
  inherited bPrecedent: TToolbarButton97
    Left = 272
  end
  inherited bSuivant: TToolbarButton97
    Left = 355
  end
  inherited bFin: TToolbarButton97
    Left = 439
  end
  inherited bAnnuler: TToolbarButton97
    Left = 189
  end
  inherited bAide: TToolbarButton97
    Left = 106
  end
  inherited GroupBox1: THGroupBox
    Width = 542
  end
  inherited P: THPageControl2
    Width = 353
    ActivePage = PDetail1
    object PDetail1: TTabSheet
      Caption = 'Fiche 1'
      ImageIndex = 1
      object lFiche1: THLabel
        Left = 4
        Top = 4
        Width = 43
        Height = 13
        Caption = 'Fiche 1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold, fsUnderline]
        ParentFont = False
      end
      object MemoDetail1: TMemo
        Left = 12
        Top = 32
        Width = 329
        Height = 205
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 0
      end
    end
    object PDetail2: TTabSheet
      Caption = 'Fiche 2'
      ImageIndex = 2
      object lFiche2: THLabel
        Left = 4
        Top = 4
        Width = 43
        Height = 13
        Caption = 'Fiche 2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold, fsUnderline]
        ParentFont = False
      end
      object MemoDetail2: TMemo
        Left = 12
        Top = 32
        Width = 329
        Height = 205
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 0
      end
    end
  end
end

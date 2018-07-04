object FAdrPiece: TFAdrPiece
  Left = 178
  Top = 106
  Width = 699
  Height = 511
  Caption = 'FAdrPiece'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 281
    Top = 2
    Width = 133
    Height = 16
    Hint = 'Nouvelle Adresse'
    Alignment = taCenter
    Caption = 'Liste des adresses'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 283
    Top = 228
    Width = 129
    Height = 16
    Caption = 'Lignes de la pi'#232'ce'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object G_LIG: THGrid
    Tag = 1
    Left = 0
    Top = 248
    Width = 687
    Height = 197
    DefaultColWidth = 30
    DefaultRowHeight = 16
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    TabOrder = 2
    OnDblClick = G_LIGDblClick
    SortedCol = -1
    Couleur = True
    MultiSelect = True
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
  end
  object G_ADR: THGrid
    Tag = 1
    Left = 0
    Top = 22
    Width = 687
    Height = 197
    DefaultRowHeight = 16
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold, fsItalic]
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goRowSelect]
    ParentFont = False
    TabOrder = 0
    OnClick = G_ADRClick
    SortedCol = 0
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
  end
  object bNouveau: THBitBtn
    Left = 652
    Top = 220
    Width = 28
    Height = 27
    Hint = 'Nouvelle Adresse'
    TabOrder = 1
    OnClick = bNouveauClick
    Margin = 1
    GlobalIndexImage = 'Z1490_S16G1'
  end
  object bValider: THBitBtn
    Left = 620
    Top = 446
    Width = 28
    Height = 27
    Hint = 'Valider'
    TabOrder = 3
    OnClick = bValiderClick
    Margin = 4
    NumGlyphs = 2
    GlobalIndexImage = 'Z0003_S16G2'
  end
  object bAnnuler: THBitBtn
    Left = 652
    Top = 446
    Width = 28
    Height = 27
    Hint = 'Annuler'
    TabOrder = 4
    OnClick = bAnnulerClick
    Margin = 3
    GlobalIndexImage = 'Z0021_S16G1'
  end
end

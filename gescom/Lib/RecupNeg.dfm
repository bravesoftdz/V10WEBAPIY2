object FRecupNeg: TFRecupNeg
  Left = 99
  Top = 160
  Width = 632
  Height = 402
  Caption = 'Traitement de r'#233'cup'#233'ration'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object TREPERTOIRE: TLabel
    Left = 5
    Top = 4
    Width = 89
    Height = 13
    Caption = 'Tables '#224' r'#233'cup'#233'rer'
  end
  object Label1: TLabel
    Left = 196
    Top = 5
    Width = 70
    Height = 13
    Caption = 'Tables '#224' traiter'
  end
  object Label2: TLabel
    Left = 356
    Top = 5
    Width = 74
    Height = 13
    Caption = 'Fichiers '#224' traiter'
    Visible = False
  end
  object Label3: TLabel
    Left = 457
    Top = 4
    Width = 39
    Height = 13
    Caption = 'Label3'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label4: TLabel
    Left = 505
    Top = 4
    Width = 39
    Height = 13
    Caption = 'Label4'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object BRecup: TButton
    Left = 519
    Top = 24
    Width = 91
    Height = 25
    Hint = 'Lancer la r'#233'cuperation compl'#232'te'
    Caption = 'Recuperation'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnClick = BRecupClick
  end
  object bVisu: TButton
    Left = 519
    Top = 125
    Width = 91
    Height = 25
    Hint = 'Voir le contenu du fichier s'#233'lectionn'#233
    Caption = 'Visu'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = bVisuClick
  end
  object HGrid1: THGrid
    Left = 4
    Top = 184
    Width = 617
    Height = 185
    DefaultRowHeight = 16
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    TabOrder = 2
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = HGrid1RowEnter
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    ColWidths = (
      151
      301
      262
      64
      64)
  end
  object ListePresent: TListBox
    Left = 5
    Top = 24
    Width = 157
    Height = 149
    ItemHeight = 13
    MultiSelect = True
    Sorted = True
    TabOrder = 3
    OnDblClick = ListePresentDblClick
  end
  object ListeATraiter: TListBox
    Left = 195
    Top = 24
    Width = 157
    Height = 149
    ItemHeight = 13
    MultiSelect = True
    Sorted = True
    TabOrder = 4
    OnDblClick = ListeATraiterDblClick
  end
  object bAjouter: THBitBtn
    Left = 166
    Top = 32
    Width = 25
    Height = 22
    Hint = 'Ajouter le fichier dans la liste '#224' traiter'
    TabOrder = 5
    OnClick = bAjouterClick
    Margin = 2
    NumGlyphs = 2
    GlobalIndexImage = 'Z0056_S16G2'
  end
  object bEnlever: THBitBtn
    Left = 166
    Top = 54
    Width = 25
    Height = 22
    Hint = 'Enlever le fichier de la liste '#224' traiter'
    TabOrder = 6
    OnClick = bEnleverClick
    Margin = 2
    NumGlyphs = 2
    GlobalIndexImage = 'Z0077_S16G2'
  end
  object bTest: TButton
    Left = 519
    Top = 49
    Width = 91
    Height = 25
    Hint = 'Lancer la r'#233'cuperation de 50 enregistrements'
    Caption = 'Test'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 7
    OnClick = bTestClick
  end
  object ListeAsso: TListBox
    Left = 28
    Top = 236
    Width = 121
    Height = 125
    ItemHeight = 13
    TabOrder = 8
    Visible = False
  end
  object ListeAssoCh: TListBox
    Left = 32
    Top = 240
    Width = 121
    Height = 125
    ItemHeight = 13
    TabOrder = 9
    Visible = False
  end
  object ListeAssoStruct: TListBox
    Left = 36
    Top = 244
    Width = 121
    Height = 125
    ItemHeight = 13
    TabOrder = 10
    Visible = False
  end
  object StringGrid1: TStringGrid
    Left = 168
    Top = 204
    Width = 281
    Height = 45
    ColCount = 50
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    TabOrder = 11
    Visible = False
  end
  object StringGrid2: TStringGrid
    Left = 171
    Top = 207
    Width = 281
    Height = 45
    ColCount = 50
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    TabOrder = 12
    Visible = False
  end
  object StringGridAsso: TStringGrid
    Left = 174
    Top = 210
    Width = 281
    Height = 77
    ColCount = 50
    FixedCols = 0
    RowCount = 500
    FixedRows = 0
    TabOrder = 13
    Visible = False
  end
  object StringGridAssoCh: TStringGrid
    Left = 177
    Top = 213
    Width = 277
    Height = 81
    ColCount = 50
    FixedCols = 0
    RowCount = 500
    FixedRows = 0
    TabOrder = 14
    Visible = False
  end
  object StringGridAssoStruct: TStringGrid
    Left = 180
    Top = 216
    Width = 281
    Height = 77
    ColCount = 50
    FixedCols = 0
    RowCount = 500
    FixedRows = 0
    TabOrder = 15
    Visible = False
  end
  object Repertoire: TEdit
    Left = 88
    Top = 328
    Width = 121
    Height = 21
    TabOrder = 16
    Text = 'Repertoire'
    Visible = False
  end
  object bStop: THBitBtn
    Left = 542
    Top = 77
    Width = 44
    Height = 44
    Hint = 'Arreter le traitement en cours'
    Caption = 'bStop'
    TabOrder = 17
    OnClick = bStopClick
    Margin = 0
    NumGlyphs = 2
    GlobalIndexImage = 'Z0461_S32G2'
  end
  object bAjouterTout: THBitBtn
    Left = 167
    Top = 100
    Width = 25
    Height = 22
    Hint = 'Tout ajouter'
    Caption = 'bAjouterTout'
    TabOrder = 18
    OnClick = bAjouterToutClick
    Margin = 2
    GlobalIndexImage = 'Z0104_S16G1'
  end
  object bEnleverTout: THBitBtn
    Left = 167
    Top = 122
    Width = 25
    Height = 22
    Hint = 'Tout enlever'
    Caption = 'BitBtn4'
    TabOrder = 19
    OnClick = bEnleverToutClick
    Margin = 2
    GlobalIndexImage = 'Z0089_S16G1'
  end
  object StringGridDECHAMPS: TStringGrid
    Left = 184
    Top = 220
    Width = 289
    Height = 57
    ColCount = 4
    FixedCols = 0
    RowCount = 500
    FixedRows = 0
    TabOrder = 20
    Visible = False
  end
  object StringGrid4: TStringGrid
    Left = 464
    Top = 212
    Width = 70
    Height = 29
    ColCount = 1
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    TabOrder = 21
    Visible = False
  end
  object bLog: TButton
    Left = 519
    Top = 150
    Width = 91
    Height = 25
    Hint = 'Voir les fichiers de log du traitement'
    Caption = 'Voir Log'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 22
    OnClick = bLogClick
  end
  object StringGridFichier: TStringGrid
    Left = 188
    Top = 224
    Width = 218
    Height = 54
    ColCount = 3
    FixedCols = 0
    RowCount = 500
    FixedRows = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    TabOrder = 23
    Visible = False
  end
  object Edit1: TEdit
    Left = 460
    Top = 180
    Width = 121
    Height = 21
    TabOrder = 24
    Visible = False
  end
  object RadioGroup1: TRadioGroup
    Left = 40
    Top = 376
    Width = 185
    Height = 37
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'Avec MAJ'
      'Sans MAJ')
    TabOrder = 25
    Visible = False
  end
  object Memo1: TMemo
    Left = 512
    Top = 264
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo1')
    TabOrder = 26
    Visible = False
  end
  object ListeATRacine: TListBox
    Left = 199
    Top = 28
    Width = 157
    Height = 149
    ItemHeight = 13
    MultiSelect = True
    Sorted = True
    TabOrder = 27
    Visible = False
    OnDblClick = ListeATraiterDblClick
  end
  object ListeFichiersSeq: TListBox
    Left = 355
    Top = 24
    Width = 157
    Height = 149
    ItemHeight = 13
    MultiSelect = True
    Sorted = True
    TabOrder = 28
    Visible = False
    OnDblClick = ListeATraiterDblClick
  end
  object RadioGroup3: TRadioGroup
    Left = 356
    Top = 19
    Width = 157
    Height = 78
    Caption = 'Nombre d'#39'enreg. de test'
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      '50'
      '500'
      '1000'
      '5000')
    TabOrder = 29
  end
  object WorkingPath: TEdit
    Left = 288
    Top = 4
    Width = 121
    Height = 21
    TabOrder = 30
    Visible = False
  end
  object RGSortie: TRadioGroup
    Left = 356
    Top = 99
    Width = 157
    Height = 78
    Caption = 'Type de traitement'
    ItemIndex = 0
    Items.Strings = (
      'Insert direct dans base'
      'Fichier texte pour SQL 7')
    TabOrder = 31
  end
  object ListeAssoCle: TListBox
    Left = 44
    Top = 248
    Width = 125
    Height = 117
    ItemHeight = 13
    TabOrder = 32
    Visible = False
  end
  object OpenDialogRep: TOpenDialog
    FileName = 'neg*'
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 596
    Top = 88
  end
end

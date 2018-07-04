object FRapsuppr: TFRapsuppr
  Left = 320
  Top = 242
  Width = 600
  Height = 240
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Rapport de suppression'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object FListe2: THGrid
    Left = 0
    Top = 0
    Width = 584
    Height = 166
    Align = alClient
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
    TabOrder = 3
    Visible = False
    OnDblClick = FListe2DblClick
    SortedCol = -1
    Titres.Strings = (
      'Code;C;S'
      'Auxiliaire;C;S'
      'G'#233'n'#233'ral;C;S'
      'Etat lettrage;C;S'
      'Remarques;L;S')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    ColWidths = (
      55
      106
      115
      90
      516)
  end
  object FListe: THGrid
    Left = 0
    Top = 0
    Width = 584
    Height = 166
    Align = alClient
    ColCount = 3
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
    TabOrder = 1
    OnDblClick = FListeDblClick
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Libell'#233
      'Rapport de suppression')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ColWidths = (
      97
      194
      278)
  end
  object Pbouton: TPanel
    Left = 0
    Top = 166
    Width = 584
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object TNbError: TLabel
      Left = 134
      Top = 11
      Width = 323
      Height = 13
      AutoSize = False
      Caption = 'TNbError'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsItalic]
      ParentFont = False
      Visible = False
      WordWrap = True
    end
    object Panel1: TPanel
      Left = 451
      Top = 2
      Width = 131
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BAide: THBitBtn
        Left = 99
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        Visible = False
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 69
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object BValider: THBitBtn
        Left = 37
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BValiderClick
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
          A400000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          0303030303030303030303030303030303030303030303030303030303030303
          03030303030303030303030303030303030303030303FF030303030303030303
          03030303030303040403030303030303030303030303030303F8F8FF03030303
          03030303030303030303040202040303030303030303030303030303F80303F8
          FF030303030303030303030303040202020204030303030303030303030303F8
          03030303F8FF0303030303030303030304020202020202040303030303030303
          0303F8030303030303F8FF030303030303030304020202FA0202020204030303
          0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
          040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
          03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
          FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
          0303030303030303030303FA0202020403030303030303030303030303F8FF03
          03F8FF03030303030303030303030303FA020202040303030303030303030303
          0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
          03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
          030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
          0202040303030303030303030303030303F8FF03F8FF03030303030303030303
          03030303FA0202030303030303030303030303030303F8FFF803030303030303
          030303030303030303FA0303030303030303030303030303030303F803030303
          0303030303030303030303030303030303030303030303030303030303030303
          0303}
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        IsControl = True
      end
      object BImprimer: THBitBtn
        Left = 5
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 131
      Height = 31
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 1
      object BRepar: THBitBtn
        Left = 63
        Top = 2
        Width = 28
        Height = 27
        Hint = 'R'#233'paration des erreurs'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BReparClick
        Margin = 2
        GlobalIndexImage = 'Z2193_S16G1'
      end
      object BRecto: THBitBtn
        Left = 33
        Top = 2
        Width = 28
        Height = 27
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BRectoClick
        Margin = 2
        GlobalIndexImage = 'Z0542_S16G1'
      end
      object BRechercher: THBitBtn
        Left = 2
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BRechercherClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BMenuZoom: THBitBtn
        Tag = -100
        Left = 94
        Top = 2
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BMenuZoomClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0016_S16G1'
        IsControl = True
      end
    end
  end
  object InfoExo: TComboBox
    Left = 208
    Top = 56
    Width = 61
    Height = 21
    Color = clYellow
    ItemHeight = 13
    TabOrder = 2
    Visible = False
  end
  object ZGen: THBitBtn
    Tag = 100
    Left = 174
    Top = 127
    Width = 33
    Height = 27
    Hint = 'Zoom sur compte g'#233'n'#233'ral'
    Caption = 'ZGen'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    Visible = False
    OnClick = ZGenClick
    NumGlyphs = 2
  end
  object ZAux: THBitBtn
    Tag = 100
    Left = 210
    Top = 127
    Width = 33
    Height = 27
    Hint = 'Zoom sur compte auxiliaire'
    Caption = 'ZAux'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = ZAuxClick
    NumGlyphs = 2
  end
  object ZAna: THBitBtn
    Tag = 100
    Left = 246
    Top = 127
    Width = 33
    Height = 27
    Hint = 'Zoom sur section analytique'
    Caption = 'ZAna'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Visible = False
    OnClick = ZAnaClick
    NumGlyphs = 2
  end
  object Cache: THCpteEdit
    Left = 184
    Top = 54
    Width = 17
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 7
    Text = '!!!'
    Visible = False
    ZoomTable = tzGeneral
    Vide = False
    Bourre = False
    okLocate = True
    SynJoker = False
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Rapport de calcul de scoring'
      'Rapport d'#39'importation'
      'Compte'
      'Nom du tiers'
      'Niveau'
      'Num'#233'ro'
      'Compte'
      'Explication'
      'Rapport sur les erreurs comptables'
      'Journal'
      'Ref. interne '
      'Pi'#232'ce/ligne'
      'Date'
      'Remarques'
      'Rapport sur les erreurs de lettrage'
      'Rapport sur les erreurs de compte'
      '16 ;'
      'Rapport sur les erreurs d'#39'Anouveau'
      'Libell'#233
      'Total d'#233'bit'
      'Total cr'#233'dit'
      'Somme d'#233'bit'
      'Somme cr'#233'dit '
      'Erreur(s) d'#233'tect'#233'e(s)'
      '24 ;'
      'G'#233'n'#233'ral'
      'Tiers'
      'Section'
      'Journal'
      'Code'
      'Rapport de contr'#244'le de doublons'
      'N'#176' Pi'#232'ce dans le fichier'
      'Identification dans la base comptable'
      '33;'
      'Rapport sur les comptes de caisse cr'#233'ditrices'
      'Compte cr'#233#233' par l'#39'utilisateur'
      '36,'
      'Rapport sur les erreurs d'#39'Immobilisations'
      'Jal dans le fichier'
      '39;')
    Left = 116
    Top = 68
  end
  object MsgQ: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Confirmez-vous la r'#233'paration des erreurs?;Q;YNC;N;C;')
    Left = 288
    Top = 59
  end
  object HMTrad: THSystemMenu
    ActiveResize = False
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 12
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 384
    Top = 68
  end
  object TimerBatch: TTimer
    Enabled = False
    Interval = 3000
    OnTimer = TimerBatchTimer
    Left = 44
    Top = 124
  end
  object PopZ: TPopupMenu
    Left = 514
    Top = 115
  end
end

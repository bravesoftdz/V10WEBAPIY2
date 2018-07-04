object FAlerteAgenda: TFAlerteAgenda
  Left = 234
  Top = 188
  BorderIcons = []
  BorderStyle = bsDialog
  Caption = 'Alerte de votre agenda'
  ClientHeight = 322
  ClientWidth = 520
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  OnDeactivate = FormDeactivate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LEnCours: TLabel
    Left = 8
    Top = 133
    Width = 81
    Height = 13
    Caption = 'Activit'#233'(s) '#224' venir'
  end
  object LRetard: TLabel
    Left = 8
    Top = 13
    Width = 91
    Height = 13
    Caption = 'Activit'#233'(s) en retard'
  end
  object SBNoAlerteRetard: THSpeedButton
    Left = 352
    Top = 8
    Width = 25
    Height = 25
    Hint = 'D'#233'sactiver les alertes s'#233'lectionn'#233'es'
    Flat = True
    ParentShowHint = False
    ShowHint = True
    OnClick = SBNoAlerteRetardClick
    GlobalIndexImage = 'Z1101_S16G1'
  end
  object SBFinRetard: THSpeedButton
    Left = 384
    Top = 8
    Width = 25
    Height = 25
    Hint = 'Consid'#233'rer les retards s'#233'lectionn'#233's comme r'#233'alis'#233's'
    Flat = True
    ParentShowHint = False
    ShowHint = True
    OnClick = SBFinRetardClick
    GlobalIndexImage = 'Z1341_S16G1'
  end
  object SBNoAlerteEnCours: THSpeedButton
    Left = 144
    Top = 128
    Width = 25
    Height = 25
    Hint = 'D'#233'sactiver les alertes s'#233'lectionn'#233'es'
    Flat = True
    ParentShowHint = False
    ShowHint = True
    OnClick = SBNoAlerteEnCoursClick
    GlobalIndexImage = 'Z1101_S16G1'
  end
  object SBReportAlerteEnCours: THSpeedButton
    Left = 176
    Top = 128
    Width = 25
    Height = 25
    Hint = 
      'Reporter les alertes s'#233'l'#233'ctionn'#233'es de la dur'#233'e sp'#233'cifi'#233'e dans "D' +
      #233'lai de report"'
    Flat = True
    ParentShowHint = False
    ShowHint = True
    OnClick = SBReportAlerteEnCoursClick
    GlobalIndexImage = 'Z2007_S16G1'
  end
  object LDelai: TLabel
    Left = 208
    Top = 134
    Width = 69
    Height = 13
    Caption = 'D'#233'lai de report'
  end
  object LBEnCours: TListBox
    Left = 9
    Top = 152
    Width = 400
    Height = 161
    Style = lbOwnerDrawFixed
    ItemHeight = 13
    MultiSelect = True
    TabOrder = 0
    OnDblClick = LBEnCoursDblClick
    OnDrawItem = LBEnCoursDrawItem
    OnKeyDown = LBEnCoursKeyDown
  end
  object BFermer: TButton
    Left = 424
    Top = 256
    Width = 89
    Height = 25
    Hint = 'Fermer la fen'#234'tre d'#39'alerte'
    Cancel = True
    Caption = 'Fermer'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    OnClick = BFermerClick
  end
  object BNOALERT: TButton
    Left = 424
    Top = 288
    Width = 89
    Height = 25
    Hint = 'Fermer la fen'#234'tre et ne plus alerter pendant cette session'
    Caption = 'Ne plus alerter'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    OnClick = BNOALERTClick
  end
  object LBRetard: TListBox
    Left = 9
    Top = 32
    Width = 400
    Height = 81
    Style = lbOwnerDrawFixed
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    MultiSelect = True
    ParentFont = False
    TabOrder = 3
    OnDblClick = LBRetardDblClick
    OnDrawItem = LBRetardDrawItem
    OnKeyDown = LBRetardKeyDown
  end
  object hcbDelaiReport: THValComboBox
    Left = 280
    Top = 130
    Width = 129
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 4
    TagDispatch = 0
    DataType = 'RTRAPPELS'
  end
  object TIAgenda: TTrayIcon
    Icon.Data = {
      0000010001002020100000000000E80200001600000028000000200000004000
      0000010004000000000080020000000000000000000000000000000000000000
      0000000080000080000000808000800000008000800080800000C0C0C0008080
      80000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000
      0000000000000000000000000000000000000000000000000000000000000008
      8888888888888888888888888300033333333333333333333333333338300877
      7777777777777777777777888330333333333333333333333333333383308BBB
      BBBBBBBBBBBBBBBBBBBFBBB3883008BBFFBBBFFBFFFBBBFBBBBBBBBB383008BF
      BBBBBFBBFBFBFBBBFFBBBFBF383008BBFBBBBBBFFFFFBBFBBBBFBBBF8330008B
      BBBBFBBBBBBBBBFBBFBFBBFFB330008BBFBBFBBFBFBBBBBFBBBBBBFFB330008B
      BBBFBBBBBFBBFFBBBFBBBFBBB330008BFBBBBFBBBBBBBFBBBFBBBBBFB330008B
      BBFBBFBFFBFFBBFBFBBBFBBFB330008BBBBBBBFBBBBBBBBBBBBBFBBBB330008B
      BFFBFBBBBBFBFBBFBBFBBBFBB330008BFBFBFB55555BBFBFBBFFBBFFB330008B
      BBBFBB59995BBFBBBBBBBBBBB330008BBBBFFB59995BFFBBBFBBBBFBF330008B
      FFBBBB59995BBFBBFFBBFBBBB330008BBBFBBB55555BBBBBBBBBBFBBF330008B
      FBBBFBBBBFFBBFBBFFBBBBBFB330008BBBFBFBBFBFBBBBFBBFBFBBBBB330008B
      BBBBBBBBBBBBBBBBBBBBBBBBB3300044444444444444444444444444443000CC
      CCCCCCCCCCCCCCCCCCCCCCCCCC3000CCCCCCCCCCCCCCCCCCCCCCCCCCCC3000CC
      CCCCCCCCCCCCCCCCCCCCCCCCCC30006000000000000000000000000000400006
      666666666666666666666666666000000000000000000000000000000000FFFF
      FFFFE0000003E000000180000000800000000000000000000000800000008000
      000080000000C0000000C0000000C0000000C0000000C0000000C0000000C000
      0000C0000000C0000000C0000000C0000000C0000000C0000000C0000000C000
      0000C0000000C0000000C0000000C0000000C0000000E0000000FFFFFFFF}
    Enabled = True
    Hint = 'Alerte agenda Cegid PGI'
    AutoShow = True
    OnLeftBtnDblClick = TIAgendaLeftBtnDblClick
    Left = 472
    Top = 40
  end
end

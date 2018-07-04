object F_Assistant: TF_Assistant
  Left = 292
  Top = 308
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Assistants '#233'tant intervenus sur la mission'
  ClientHeight = 270
  ClientWidth = 625
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GrilleAssistant: THGrid
    Left = 0
    Top = 0
    Width = 625
    Height = 270
    Align = alClient
    DefaultRowHeight = 18
    Enabled = False
    FixedCols = 0
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    ParentFont = False
    ScrollBars = ssVertical
    TabOrder = 0
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      64
      64
      64
      64
      64)
  end
end

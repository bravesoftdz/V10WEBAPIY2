object FVisuLog: TFVisuLog
  Left = 379
  Top = 107
  Width = 696
  Height = 480
  Caption = 'R'#233'sultats de l'#39'import'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ListBox1: TListBox
    Left = 8
    Top = 16
    Width = 121
    Height = 330
    ItemHeight = 13
    TabOrder = 0
    OnClick = ListBox1Click
  end
  object Button1: TButton
    Left = 140
    Top = 356
    Width = 75
    Height = 25
    Caption = 'Imprimer'
    TabOrder = 1
    OnClick = Button1Click
  end
  object RichEdit1: TRichEdit
    Left = 140
    Top = 16
    Width = 537
    Height = 331
    BorderWidth = 1
    Color = clBtnFace
    Ctl3D = False
    Lines.Strings = (
      'RichEdit1')
    ParentCtl3D = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object CheminLog: TEdit
    Left = 84
    Top = 392
    Width = 121
    Height = 21
    TabOrder = 3
    Visible = False
  end
  object PrintDialog1: TPrintDialog
    Copies = 1
    FromPage = 1
    MaxPage = 99
    Options = [poDisablePrintToFile]
    ToPage = 99
    Left = 36
    Top = 392
  end
end

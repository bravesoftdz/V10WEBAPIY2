object Form1: TForm1
  Left = 226
  Top = 235
  Width = 870
  Height = 640
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object HValComboBox1: THValComboBox
    Left = 416
    Top = 120
    Width = 145
    Height = 21
    ItemHeight = 13
    TabOrder = 0
    OnChange = HValComboBox1Change
    Items.Strings = (
      'LILI'
      'LOLO'
      'LALA')
    TagDispatch = 0
    Values.Strings = (
      '1'
      '2'
      '3')
  end
  object RESULTVALUE: THCritMaskEdit
    Left = 328
    Top = 192
    Width = 265
    Height = 21
    TabOrder = 1
    TagDispatch = 0
  end
  object RESULTEXT: THCritMaskEdit
    Left = 328
    Top = 224
    Width = 265
    Height = 21
    TabOrder = 2
    TagDispatch = 0
  end
end

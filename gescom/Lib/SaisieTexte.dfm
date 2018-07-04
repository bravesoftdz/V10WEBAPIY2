object FSaisieTexte: TFSaisieTexte
  Left = 386
  Top = 170
  Width = 213
  Height = 103
  Caption = 'Saisie du texte'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Edit1: TEdit
    Left = 16
    Top = 12
    Width = 173
    Height = 21
    TabOrder = 0
  end
  object Button1: TButton
    Left = 16
    Top = 44
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 114
    Top = 44
    Width = 75
    Height = 25
    Caption = 'Annuler'
    TabOrder = 2
    OnClick = Button2Click
  end
end

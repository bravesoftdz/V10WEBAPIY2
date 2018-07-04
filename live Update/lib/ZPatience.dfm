object FPatience: TFPatience
  Left = 502
  Top = 386
  BorderIcons = []
  BorderStyle = bsDialog
  Caption = 'Veuillez patienter...'
  ClientHeight = 117
  ClientWidth = 452
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object lCreation: TLabel
    Left = 22
    Top = 33
    Width = 9
    Height = 13
    Caption = '...'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object lTitre: TLabel
    Left = 14
    Top = 7
    Width = 27
    Height = 13
    Caption = 'Titre'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold, fsUnderline]
    ParentColor = False
    ParentFont = False
  end
  object lAide: TLabel
    Left = 21
    Top = 37
    Width = 417
    Height = 14
    AutoSize = False
    Caption = '...'
  end
  object ProgressBar1: TProgressBar
    Left = 6
    Top = 67
    Width = 440
    Height = 17
    TabOrder = 0
    Visible = False
  end
end

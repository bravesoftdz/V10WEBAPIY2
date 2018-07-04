object FPatience: TFPatience
  Left = 366
  Top = 368
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
  object lAide: THLabel
    Left = 21
    Top = 53
    Width = 417
    Height = 14
    AutoSize = False
    Caption = '...'
  end
  object lCreation: THLabel
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
  object lTitre: THLabel
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
  object sbBarre: THStatusBar
    Left = 0
    Top = 93
    Width = 452
    Height = 24
    Panels = <
      item
        Width = 172
      end
      item
        Alignment = taCenter
        Text = 'NUM'
        Width = 35
      end
      item
        Alignment = taCenter
        Width = 35
      end
      item
        Alignment = taCenter
        Text = '   '
        Width = 35
      end
      item
        Alignment = taCenter
        Text = '16:19:36'
        Width = 55
      end
      item
        Alignment = taCenter
        Text = 'mar. 11 d'#233'c. 2007'
        Width = 120
      end>
  end
  object ProgressBar1: TProgressBar
    Left = 0
    Top = 76
    Width = 452
    Height = 17
    Align = alBottom
    TabOrder = 1
    Visible = False
  end
end

object FMain: TFMain
  Left = 245
  Top = 231
  Width = 434
  Height = 131
  Caption = 'Importation automatique'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 426
    Height = 85
    Align = alClient
    TabOrder = 0
    object FlashingLabel1: TFlashingLabel
      Left = 24
      Top = 32
      Width = 364
      Height = 20
      Caption = 'Restauration en cours ... Veuillez patienter ...'
      Color = clMaroon
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Flashing = True
    end
  end
  object Status: THStatusBar
    Left = 0
    Top = 85
    Width = 426
    Height = 19
    Panels = <
      item
        Width = 181
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
        Text = '16:41:11'
        Width = 55
      end
      item
        Alignment = taCenter
        Text = 'mer. 25 oct. 2000'
        Width = 120
      end>
    SimplePanel = False
  end
  object Timer: TTimer
    Interval = 2000
    OnTimer = TimerTimer
    Left = 280
    Top = 24
  end
end

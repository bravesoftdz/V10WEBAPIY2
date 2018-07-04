object Form1: TForm1
  Left = 50
  Top = 205
  Width = 511
  Height = 653
  Caption = 'Services'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object ListView1: TListView
    Left = 0
    Top = 0
    Width = 503
    Height = 599
    Align = alClient
    Columns = <
      item
        Caption = 'Nom'
        Width = 225
      end
      item
        Caption = 'Nom interne'
        Width = 150
      end
      item
        Caption = 'Etat'
        Width = 100
      end>
    MultiSelect = True
    TabOrder = 0
    ViewStyle = vsReport
  end
  object Memo1: TMemo
    Left = 136
    Top = 32
    Width = 113
    Height = 569
    Lines.Strings = (
      'Memo1'
      'Destin'#233' '#224' accueillir '
      'les noms des services'
      ' dans le m'#234'me ordre'
      ' que ListView1 pour'
      ' les retrouver plus '
      'vite.')
    TabOrder = 1
    Visible = False
    WordWrap = False
  end
  object MainMenu1: TMainMenu
    Left = 32
    Top = 40
    object Start1: TMenuItem
      Caption = 'Start'
      OnClick = ChgEtatService
    end
    object Stop1: TMenuItem
      Caption = 'Stop'
      OnClick = ChgEtatService
    end
    object Pause1: TMenuItem
      Caption = 'Pause'
      OnClick = ChgEtatService
    end
    object Continue1: TMenuItem
      Caption = 'Continue'
      OnClick = ChgEtatService
    end
    object Refresh1: TMenuItem
      Caption = 'Refresh'
      ShortCut = 116
      OnClick = Rafraichir
    end
    object Recreate1: TMenuItem
      Caption = 'Recreate'
      OnClick = Recreate1Click
    end
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 5000
    OnTimer = Rafraichir
    Left = 32
    Top = 8
  end
end

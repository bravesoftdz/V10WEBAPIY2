object FicheInfoToxZip: TFicheInfoToxZip
  Left = 371
  Top = 151
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Caractéristiques du fichier :'
  ClientHeight = 113
  ClientWidth = 332
  Color = clInfoBk
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object LBefore: TLabel
    Left = 8
    Top = 8
    Width = 208
    Height = 16
    Caption = 'Taille en octets avant compression :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object LAfter: TLabel
    Left = 8
    Top = 36
    Width = 210
    Height = 16
    Caption = 'Taille en octets après compression :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object LRate: TLabel
    Left = 8
    Top = 64
    Width = 132
    Height = 16
    Caption = 'Taux de compression :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object LDate: TLabel
    Left = 8
    Top = 92
    Width = 135
    Height = 16
    Caption = 'Dernière modification le'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object LTime: TLabel
    Left = 242
    Top = 92
    Width = 7
    Height = 16
    Caption = 'à'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object EBfore: TLabel
    Left = 239
    Top = 8
    Width = 82
    Height = 16
    Alignment = taRightJustify
    Caption = '0,000,000,000'
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object EAfter: TLabel
    Left = 239
    Top = 36
    Width = 82
    Height = 16
    Alignment = taRightJustify
    Caption = '0,000,000,000'
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ERate: TStaticText
    Left = 147
    Top = 64
    Width = 29
    Height = 20
    Caption = '0,00'
    Color = clInfoBk
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    TabOrder = 0
  end
  object EDate: TStaticText
    Left = 153
    Top = 92
    Width = 68
    Height = 20
    Caption = '01/01/1900'
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
  end
  object ETime: TStaticText
    Left = 261
    Top = 92
    Width = 54
    Height = 20
    Caption = '00:00:00'
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
  end
end
